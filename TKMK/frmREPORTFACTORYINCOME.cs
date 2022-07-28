using System;
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
                                    ,(SELECT ROUND(ISNULL(SUM([SALESMMONEYS]),0),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001) AS 'GROUPMONEYS'
                                    ,(SUM(TA026)-(SELECT ROUND(ISNULL(SUM([SALESMMONEYS]),0),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001)) AS 'VISITORMONEYS'
                                    ,(SELECT ISNULL(SUM(CARNUM),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001) AS 'CARNUM'
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
                            [INDATES] AS '日期',[YEARS] AS '年',[WEEKS] AS '週',[TOTALMONEYS] AS 銷售組當日業績,[GROUPMONEYS] AS '團客業績',([TOTALMONEYS]-[GROUPMONEYS]) AS '散客業績',[CARNUM] AS '遊覽車次',[CARAVGMONEYS] AS '每車平均業績'
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
                            ,(SUM([TOTALMONEYS])-SUM([GROUPMONEYS])) AS '散客業績'
                            ,SUM([CARNUM]) AS '遊覽車次'
                            ,AVG([CARAVGMONEYS]) AS '每車平均業績'
                            ,(SELECT SUM(TOTALMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')   AS '同期業績'
                            ,(SELECT SUM(GROUPMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期團客'
                            ,(SELECT (SUM([TOTALMONEYS])-SUM([GROUPMONEYS])) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期散客'
                            ,(SELECT SUM(CARNUM) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期車次'
                            ,(SELECT AVG(CARAVGMONEYS) FROM [TKMK].[dbo].[TBFACTORYINCOME] LASTTBFACTORYINCOME WHERE LASTTBFACTORYINCOME.WEEKS=[TBFACTORYINCOME].WEEKS AND LASTTBFACTORYINCOME.YEARS='{1}')  AS '同期平均業績'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE YEARS='{0}'
                            AND WEEKS IN ('{2}','{3}' ,'{4}' ,'{5}')
                            GROUP BY [YEARS],[WEEKS]
 

                            ", YEARS, LASTYEARS, WEEKS1, WEEKS2, WEEKS3, WEEKS4);

            return SB;

        }


        public void SETFASTREPORT3(string YEARS)
        {
            StringBuilder SQL4 = new StringBuilder();

            SQL4 = SETSQL4(YEARS);
         

            Report report1 = new Report();
            report1.Load(@"REPORT\觀光賣場來車和金額.frx");

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
            table.SelectCommand = SQL4.ToString();


            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL4(string YEARS)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT YEAR(CONVERT(DATETIME,INDATES)) AS '年度',MONTH(CONVERT(DATETIME,INDATES)) AS '月份',SUM([TOTALMONEYS]) '賣場總金額',SUM([GROUPMONEYS]) '團客金額',(SUM([TOTALMONEYS])-SUM([GROUPMONEYS])) '散客金額',SUM([CARNUM]) '來車數'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE INDATES LIKE '{0}%'
                            GROUP BY YEAR(CONVERT(DATETIME,INDATES)),MONTH(CONVERT(DATETIME,INDATES))
 

                            ", YEARS );

            return SB;

        }

        public void SETFASTREPORT4(string SDATE,string EDATES)
        {
            StringBuilder SQL4 = new StringBuilder();

            SQL4.Clear();
            SQL4.AppendFormat(@"
                                --20220728 查團車
                                SELECT  
                                CONVERT(nvarchar,[CREATEDATES],112) AS '日期'
                                ,[SERNO] AS '序號'
                                ,[CARNAME] AS '車名'
                                ,[CARNO] AS '車號'
                                ,[CARKIND] AS '車種'
                                ,[GROUPKIND]  AS '團類'
                                ,[ISEXCHANGE] AS '兌換券'
                                ,[EXCHANGETOTALMONEYS] AS '券總額'
                                ,[EXCHANGESALESMMONEYS] AS '券消費'
                                ,[SALESMMONEYS] AS '消費總額'
                                ,[SPECIALMNUMS] AS '特賣數'
                                ,[SPECIALMONEYS] AS '特賣獎金'
                                ,[COMMISSIONBASEMONEYS] AS '茶水費'
                                ,[COMMISSIONPCTMONEYS] AS '消費獎金'
                                ,[TOTALCOMMISSIONMONEYS] AS '總獎金'
                                ,[CARNUM] AS '車數'
                                ,[GUSETNUM] AS '來客數'
                                ,[EXCHANNO] AS '優惠券名'
                                ,[EXCHANACOOUNT] AS '優惠券帳號'
                                ,CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'
                                ,CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間'
                                ,[STATUS] AS '狀態'
                                ,CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間'
                                ,CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間'
                                ,[EXCHANGEMONEYS] AS '領券額'
                                ,[ID]
                                ,[CREATEDATES]
                                FROM [TKMK].[dbo].[GROUPSALES]
                                WHERE CONVERT(nvarchar,[CREATEDATES],112)>='{0}' AND CONVERT(nvarchar,[CREATEDATES],112)<='{1}'
                                AND [STATUS]<>'取消預約'
                                ORDER BY CONVERT(nvarchar,[CREATEDATES],112),SERNO

                                ", SDATE,EDATES);


             Report report1 = new Report();
            report1.Load(@"REPORT\團車明細.frx");

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
            table.SelectCommand = SQL4.ToString();


            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
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
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3(dateTimePicker3.Value.ToString("yyyy"));
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
