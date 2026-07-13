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
            dateTimePicker9.Value = MONTHFirstDay;
            dateTimePicker10.Value = LastDay;
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

        public void SETFASTREPORT5(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2= new StringBuilder();

            SQL1 = SETSQL5A(SDATES, EDATES);
            SQL2 = SETSQL5B(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\團車旅遊圖表.frx");

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
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();


            report1.SetParameterValue("P1", SDATES);
            report1.SetParameterValue("P2", EDATES);

            report1.Preview = previewControl5;
            report1.Show();
        }

        public StringBuilder SETSQL5A(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                              
                             SELECT 
                                [PLAYDAYKINDS] + ' ' + CONVERT(NVARCHAR, CAST(COUNT([PLAYDAYKINDS]) * 100.0 / SUM(COUNT([PLAYDAYKINDS])) OVER () AS DECIMAL(5, 2))) + '%' AS [PLAYDAYKINDS],
                                COUNT([PLAYDAYKINDS]) AS NUM
                            FROM 
                                [TKMK].[dbo].[GROUPSALES]
                            WHERE 
                                CONVERT(NVARCHAR,[CREATEDATES],112) >= '{0}'
	                            AND  CONVERT(NVARCHAR,[CREATEDATES],112) <= '{1}'
                            GROUP BY 
                                [PLAYDAYKINDS]
                            ORDER BY 
                                COUNT([PLAYDAYKINDS]) DESC

                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQL5B(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 
                                [PLAYDAYS] + ' ' + CONVERT(NVARCHAR, CAST(COUNT([PLAYDAYS]) * 100.0 / SUM(COUNT([PLAYDAYS])) OVER () AS DECIMAL(5, 2))) + '%' AS [PLAYDAYS],
                                COUNT([PLAYDAYS]) AS NUM
                            FROM 
                                [TKMK].[dbo].[GROUPSALES]
                            WHERE 
                                CONVERT(NVARCHAR,[CREATEDATES],112) >= '{0}'
	                            AND  CONVERT(NVARCHAR,[CREATEDATES],112) <= '{1}'
                            GROUP BY 
                                [PLAYDAYS]
                            ORDER BY 
                                COUNT([PLAYDAYS]) DESC                             

                            ", SDATES, EDATES);

            return SB;

        }

        public void SETFASTREPORT6(string SDATES, string EDATES,string CARNO)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL6(SDATES, EDATES,CARNO);
           
            Report report1 = new Report();
            report1.Load(@"REPORT\團車遊程次數圖表.frx");

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

            report1.Preview = previewControl6;
            report1.Show();
        }

        public StringBuilder SETSQL6(string SDATES, string EDATES,string CARNO)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();

            if (!string.IsNullOrEmpty(CARNO))
            {
                QUERY1.AppendFormat(@" AND [CARNO] LIKE '%{0}%' ", CARNO);
            }
            else
            {
                QUERY1.AppendFormat(@"  ");
            }


            SB.AppendFormat(@"                              
                             SELECT  
                            [CARNAME] AS '車名'
                            ,[CARNO] AS '車號'
                            ,[PLAYPROCESS] AS '去程回程'   
                            ,COUNT([CARNO])  AS '遊程次數'
                            ,AVG([SALESMMONEYS]) AS '平均消費金額'
                            ,SUM([SALESMMONEYS]) AS '加總消費金額'

                            FROM [TKMK].[dbo].[GROUPSALES] 
                            WHERE 1=1
                            AND CONVERT(nvarchar,[CREATEDATES],112)>='{0}'
                            AND CONVERT(nvarchar,[CREATEDATES],112)<='{1}'
                            AND [STATUS]<>'取消預約'
                            {2}
                            GROUP BY [CARNAME],[CARNO],[PLAYPROCESS] 
                            ORDER BY COUNT([CARNO])  DESC,[CARNAME],[CARNO],[PLAYPROCESS] 

                            ", SDATES, EDATES, QUERY1.ToString());

            return SB;

        }

        public void SETFASTREPORT7(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL7(SDATES, EDATES);

            Report report1 = new Report();
            report1.Load(@"REPORT\團車入場時間.frx");

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

            report1.Preview = previewControl7;
            report1.Show();
        }

        public StringBuilder SETSQL7(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();

        


            SB.AppendFormat(@"                              
                            SELECT 
                                 V.ID
                                ,V.SETHOURS
                                ,ISNULL(TEMP.HRS, V.SETHOURS) AS '入場時間' -- 防止 LEFT JOIN 沒資料時顯示 NULL
                                ,ISNULL(TEMP.SUMMONEYS, 0) AS '團車銷售總金額' -- 沒資料時自動補 0
                                ,ISNULL(TEMP.SUMCARNUMS, 0) AS '團車來車數'
                                ,ISNULL(TEMP.AVGMONEYS, 0) AS '團車平均銷售金額 '
                            FROM [TKMK].[dbo].[VISITORS_HOURS] AS V

                            LEFT JOIN 
                            (
                                SELECT
                                     DATEPART(HOUR, [GROUPSTARTDATES]) AS 'HRS'
                                    ,SUM([SALESMMONEYS]) AS 'SUMMONEYS'
                                    ,SUM([CARNUM]) AS 'SUMCARNUMS'
                                    -- 優化 1：加入 CASE WHEN 防止車數為 0 時引發除以零錯誤
                                    ,CASE 
                                        WHEN SUM([CARNUM]) = 0 THEN 0 
                                        ELSE SUM([SALESMMONEYS]) / SUM([CARNUM]) 
                                     END AS 'AVGMONEYS'
                                FROM [TKMK].[dbo].[GROUPSALES]
                                -- 優化 2：移除欄位上的 CONVERT 轉換，改用標準日期區間比對（支援索引）
                                WHERE CONVERT(NVARCHAR,[GROUPSTARTDATES],112) >= '{0}' 
                                  AND  CONVERT(NVARCHAR,[GROUPSTARTDATES],112)<= '{1}' 
                                GROUP BY DATEPART(HOUR, [GROUPSTARTDATES])
                            ) AS TEMP ON V.[SETHOURS] = TEMP.HRS
                            ORDER BY V.SETHOURS;

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

        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT5(dateTimePicker9.Value.ToString("yyyyMMdd"), dateTimePicker10.Value.ToString("yyyyMMdd"));
        }
        private void button6_Click(object sender, EventArgs e)
        {
            string SDATES= dateTimePicker11.Value.ToString("yyyyMMdd");
            string EDATES= dateTimePicker12.Value.ToString("yyyyMMdd");
            string CARNO= textBox1.Text.Trim();
            
            SETFASTREPORT6(SDATES, EDATES, CARNO);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker13.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker14.Value.ToString("yyyyMMdd");

            SETFASTREPORT7(SDATES, EDATES);

        }
        #endregion


    }
}
