﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Globalization;

namespace TKMK
{
    public partial class frmREPORTFACTORYINCOME : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        SqlTransaction tran;
        int result;

        public frmREPORTFACTORYINCOME()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            DateTime nowTime = DateTime.Now;
            #region 獲取本週第一天
            //星期一為第一天  
            int weeknow = Convert.ToInt32(nowTime.DayOfWeek);

            //因為是以星期一為第一天，所以要判斷weeknow等於0時，要向前推6天。  
            weeknow = (weeknow == 0 ? (7 - 1) : (weeknow - 1));
            int daydiff = (-1) * weeknow-7;

            //本週第一天  
            DateTime FirstDay = nowTime.AddDays(daydiff);

            dateTimePicker1.Value = FirstDay;
            #endregion

            #region 獲取本週最後一天
            //星期天為最後一天  
            int lastWeekDay = Convert.ToInt32(nowTime.DayOfWeek);
            lastWeekDay = lastWeekDay == 0 ? (7 - lastWeekDay) : lastWeekDay;
            int lastWeekDiff = (7 - lastWeekDay)-7;

            //本週最後一天  
            DateTime LastDay = nowTime.AddDays(lastWeekDiff);

            dateTimePicker2.Value = LastDay;
            #endregion
        }

        public void ADDTKMK_TBFACTORYINCOME(string SDATES,string EDATES)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" 
                                    DELETE [TKMK].[dbo].[TBFACTORYINCOME]
                                    WHERE INDATES>='{0}' AND INDATES<='{1}'

                                    INSERT INTO [TKMK].[dbo].[TBFACTORYINCOME]
                                    ([INDATES],[YEARS],[WEEKS],[TOTALMONEYS],[GROUPMONEYS],[VISITORMONEYS],[CARNUM],[CARAVGMONEYS])

                                    SELECT INDATES,YEARS,WEEKS,CONVERT(INT,TOTALMONEYS) TOTALMONEYS,CONVERT(INT,GROUPMONEYS)  GROUPMONEYS,CONVERT(INT,VISITORMONEYS)  VISITORMONEYS,CARNUM
                                    ,CASE WHEN CARNUM>0 THEN CONVERT(INT,ROUND(GROUPMONEYS/CARNUM,0))  ELSE 0 END AS 'CARAVGMONEYS'
                                    FROM (
                                    SELECT 
                                    TA001 AS 'INDATES'
                                    ,DATEPART(YEAR, [TA001]) AS YEARS
                                    ,DATEPART(Week, [TA001]) AS WEEKS
                                    ,SUM(TA026) AS 'TOTALMONEYS'
                                    ,(SELECT ROUND(ISNULL(SUM([SALESMMONEYS]),0)/1.05,0) FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=TA001) AS 'GROUPMONEYS'
                                    ,(SUM(TA026)-(SELECT ROUND(ISNULL(SUM([SALESMMONEYS]),0)/1.05,0) FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=TA001)) AS 'VISITORMONEYS'
                                    ,(SELECT ISNULL(SUM(CARNUM),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=TA001) AS 'CARNUM'
                                    FROM [TK].dbo.POSTA
                                    WHERE TA002 IN ('106701','106702')
                                    AND TA001>='{0}' AND TA001<='{1}'
                                    GROUP BY TA001
                                    ) AS TEMP
                                    ORDER BY INDATES
                                    ", SDATES, EDATES);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("完成");

                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\觀光業績及車次明細表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@" 
                            SELECT 
                            [INDATES] AS '日期',[YEARS] AS '年',[WEEKS] AS '週',[TOTALMONEYS] AS 銷售組當日業績,[GROUPMONEYS] AS '團客業績',[VISITORMONEYS] AS '散客業績',[CARNUM] AS '遊覽車次',[CARAVGMONEYS] AS '每車平均業績'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE [INDATES]>='{0}' AND [INDATES]<='{1}'
 

                            ", SDATES, EDATES);

            return SB;

        }

        public void SETFASTREPORT2(DateTime SDATES)
        {
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();


            string YEARS = SDATES.ToString("yyyy");
            string LASTYEARS = SDATES.AddYears(-1).ToString("yyyy");

            string INDATES = SDATES.ToString("yyyyMM");
            string LASTINDATES = SDATES.AddYears(-1).ToString("yyyyMM");

            GregorianCalendar gc = new GregorianCalendar();
            int WEEKOFYEARS = gc.GetWeekOfYear(SDATES, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            string WEEKS1 = (WEEKOFYEARS - 1).ToString();
            string WEEKS2 = (WEEKOFYEARS - 2).ToString();
            string WEEKS3 = (WEEKOFYEARS - 3).ToString();
            string WEEKS4 = (WEEKOFYEARS - 4).ToString();



            SQL2 = SETSQL2(YEARS, LASTYEARS, INDATES, LASTINDATES);
            SQL3 = SETSQL3(YEARS, LASTYEARS, WEEKS1, WEEKS2, WEEKS3, WEEKS4);

            Report report1 = new Report();
            report1.Load(@"REPORT\觀光業績及車次比較表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL2.ToString();

            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL3.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL2(string YEARS, string LASTYEARS,string INDATES,string LASTINDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT [YEARS] AS '年月',SUM([CARNUM]) AS '來車數' 
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE [YEARS]='{0}'
                            GROUP BY [YEARS]
                            UNION ALL
                            SELECT [YEARS],SUM([CARNUM]) AS 'CARNUM'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE [YEARS]='{1}'
                            GROUP BY [YEARS]
                            UNION ALL
                            SELECT SUBSTRING([INDATES],1,6),SUM([CARNUM]) AS 'CARNUM'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE[INDATES] LIKE '{2}%'
                            GROUP BY SUBSTRING([INDATES],1,6)
                            UNION ALL
                            SELECT SUBSTRING([INDATES],1,6),SUM([CARNUM]) AS 'CARNUM'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE [INDATES] LIKE '{3}%'
                            GROUP BY SUBSTRING([INDATES],1,6)
 

                            ", YEARS, LASTYEARS, INDATES, LASTINDATES);

            return SB;

        }

        public StringBuilder SETSQL3(string YEARS,string LASTYEARS,string WEEKS1, string WEEKS2, string WEEKS3, string WEEKS4)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 
                            [YEARS] AS '年度'
                            ,[WEEKS] AS '週次'
                            ,SUM([TOTALMONEYS]) AS '銷售組業績'
                            ,SUM([GROUPMONEYS]) AS '團客業績'
                            ,SUM([VISITORMONEYS]) AS '散客業績'
                            ,SUM([CARNUM]) AS '遊覽車次'
                            ,AVG([CARAVGMONEYS]) AS '每車平均業績'
                            ,(SELECT SUM(TOTALMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')   AS '同期業績'
                            ,(SELECT SUM(GROUPMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期團客'
                            ,(SELECT SUM(VISITORMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期散客'
                            ,(SELECT SUM(CARNUM) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期車次'
                            ,(SELECT AVG(CARAVGMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期平均業績'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE YEARS='{0}'
                            AND WEEKS IN ('{2}','{3}' ,'{4}' ,'{5}')
                            GROUP BY [YEARS],[WEEKS]
 

                            ", YEARS, LASTYEARS, WEEKS1, WEEKS2, WEEKS3, WEEKS4);

            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            ADDTKMK_TBFACTORYINCOME(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker1.Value);
        }

        #endregion


    }
}