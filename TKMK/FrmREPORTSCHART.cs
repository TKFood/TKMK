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
    public partial class FrmREPORTSCHART : Form
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

        public FrmREPORTSCHART()
        {
            InitializeComponent();

            SETDATES();
        }




        #region FUNCTION
        public void SETDATES()
        {
            DateTime FirstDay = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
            DateTime MONTHFirstDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = LastDay;
            dateTimePicker3.Value = MONTHFirstDay;
            dateTimePicker4.Value = LastDay;
            dateTimePicker5.Value = MONTHFirstDay;
            dateTimePicker6.Value = LastDay;
            dateTimePicker7.Value = MONTHFirstDay;
            dateTimePicker8.Value = LastDay;
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\團車業績圖表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report1.Dictionary.Connections[0].CommandTimeout = 180;


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
                            SELECT *
                            FROM (
                            SELECT
                            '全部' AS '分類'
                            ,YEAR([CREATEDATES]) AS '年'
                            ,MONTH([CREATEDATES]) AS '月份'
                            ,COUNT([CARNUM])  AS  '車數'
                            ,SUM([SALESMMONEYS])  AS  '總團車銷售金額'
                            ,SUM([COMMISSIONBASEMONEYS])  AS  '總茶水費'
                            ,SUM([COMMISSIONPCTMONEYS])      AS  '總佣金' 
                            ,SUM([TOTALCOMMISSIONMONEYS])  AS  '總佣金+總茶水費'
                            ,(
                            SELECT SUM(TA026) 
                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            WHERE TA002 LIKE '1067%'
                            AND YEAR(TA001)=YEAR([CREATEDATES]) AND MONTH(TA001)=MONTH([CREATEDATES])
                            ) AS  '觀光+硯微墨的總銷售金額'

                            FROM [TKMK].[dbo].[GROUPSALES]
                            WHERE [STATUS]='完成接團'
                            AND CONVERT(NVARCHAR,[CREATEDATES],112)>='{0}' AND CONVERT(NVARCHAR,[CREATEDATES],112)<='{1}' 
                            GROUP BY YEAR([CREATEDATES]),MONTH([CREATEDATES])
                            UNION ALL

                            SELECT
                            '滿5000元以上' AS '分類'
                            ,YEAR([CREATEDATES]) AS '年'
                            ,MONTH([CREATEDATES]) AS '月份'
                            ,COUNT([CARNUM])  AS  '車數'
                            ,SUM([SALESMMONEYS])  AS  '總團車銷售金額'
                            ,SUM([COMMISSIONBASEMONEYS])  AS  '總茶水費'
                            ,SUM([COMMISSIONPCTMONEYS])      AS  '總佣金' 
                            ,SUM([TOTALCOMMISSIONMONEYS])  AS  '總佣金+總茶水費'
                           ,(
                            SELECT SUM(TA026) 
                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            WHERE TA002 LIKE '1067%'
                            AND YEAR(TA001)=YEAR([CREATEDATES]) AND MONTH(TA001)=MONTH([CREATEDATES])
                            ) AS  '觀光+硯微墨的總銷售金額'

                            FROM [TKMK].[dbo].[GROUPSALES]
                            WHERE [STATUS]='完成接團'
                            AND CONVERT(NVARCHAR,[CREATEDATES],112)>='{0}' AND CONVERT(NVARCHAR,[CREATEDATES],112)<='{1}' 
                            AND [SALESMMONEYS]>=5000
                            GROUP BY YEAR([CREATEDATES]),MONTH([CREATEDATES])
                            ) AS TEMP
                            ORDER BY 分類,年,月份

   

                            ", SDATES, EDATES);

            return SB;

        }

        public void SETFASTREPORT2(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\團車類型圖表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report1.Dictionary.Connections[0].CommandTimeout = 180;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SDATES);
            report1.SetParameterValue("P2", EDATES);

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                              
                           SELECT 
                                [GROUPKIND] + ' ' + CONVERT(NVARCHAR, CAST(COUNT([GROUPKIND]) * 100.0 / SUM(COUNT([GROUPKIND])) OVER () AS DECIMAL(5, 2))) + '%' AS GROUPKIND,
                                COUNT([GROUPKIND]) AS NUM
                            FROM 
                                [TKMK].[dbo].[GROUPSALES]
                            WHERE 
                                CONVERT(NVARCHAR,[CREATEDATES],112) >= '{0}'
	                            AND  CONVERT(NVARCHAR,[CREATEDATES],112) <= '{1}'
                            GROUP BY 
                                [GROUPKIND]
                            ORDER BY 
                                COUNT([GROUPKIND]) DESC


   

                            ", SDATES, EDATES);

            return SB;

        }

        public void SETFASTREPORT3(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\團車類型明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report1.Dictionary.Connections[0].CommandTimeout = 180;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SDATES);
            report1.SetParameterValue("P2", EDATES);

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                              
                            SELECT 
                            [GROUPKIND],
                            COUNT([GROUPKIND]) AS '來車數',
                            SUM(SALESMMONEYS) AS '銷售總金額',
                            SUM([TOTALCOMMISSIONMONEYS]) AS '佣金總金額',
                            SUM([GUSETNUM]) AS '結帳筆數',
                            SUM(SALESMMONEYS)/COUNT([GROUPKIND]) AS '每車平均銷售金額',
                            SUM([GUSETNUM])/COUNT([GROUPKIND])  AS '每車平均結帳筆數'

                            FROM 
                                [TKMK].[dbo].[GROUPSALES]
                            WHERE 
                                CONVERT(NVARCHAR,[CREATEDATES],112) >= '{0}'
	                            AND  CONVERT(NVARCHAR,[CREATEDATES],112) <= '{1}'
                            GROUP BY 
                                [GROUPKIND]
                            ORDER BY 
                                COUNT([GROUPKIND]) DESC



   

                            ", SDATES, EDATES);

            return SB;

        }

        public void SETFASTREPORT4(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL4(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\團車商品明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report1.Dictionary.Connections[0].CommandTimeout = 180;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SDATES);
            report1.SetParameterValue("P2", EDATES);

            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL4(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                              
                            SELECT 
                            團類,
                            品號,
                            品名,
                            單位,
                            SUM(銷售數量) 銷售數量,
                            SUM(銷售未稅金額) 銷售未稅金額

                            FROM 
                            (
	                            SELECT 
	                            [GROUPKIND] AS '團類',
	                            POSTA.[TA008],
	                            TA001,TA002,TA003,TA006,
	                            TB001,TB002,TB003,TB006,
	                            TB010 AS '品號',
	                            MB002 AS '品名',
	                            MB004 AS '單位',
	                            TB019 AS '銷售數量',
	                            TB031 AS '銷售未稅金額'

	                            FROM 
		                            [TKMK].[dbo].[GROUPSALES]
		                            LEFT JOIN [TK].dbo.POSTA ON POSTA.TA008=[GROUPSALES].TA008 AND  POSTA.TA001=CONVERT(NVARCHAR,[GROUPSALES].[CREATEDATES],112)
		                            LEFT JOIN [TK].dbo.POSTB ON TB001=TA001 AND TB002=TA002 AND TB003=TA003 AND TB006=TA006
		                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
	                            WHERE 
		                            CONVERT(NVARCHAR,[CREATEDATES],112) >= '{0}'
		                            AND  CONVERT(NVARCHAR,[CREATEDATES],112) <= '{1}'
                            ) AS TEMP
                            GROUP BY 
                            團類,
                            品號,
                            品名,
                            單位
                            HAVING (SUM(銷售未稅金額))>0
                            ORDER BY 團類,SUM(銷售未稅金額) DESC             


   

                            ", SDATES, EDATES);

            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));

        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3(dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
