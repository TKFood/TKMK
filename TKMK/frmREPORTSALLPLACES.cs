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
    public partial class frmREPORTSALLPLACES : Form
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


        public frmREPORTSALLPLACES()
        {
            InitializeComponent();

            SETDATES();

        }

        #region FUNCTION
        public void SETDATES()
        {
            DateTime FirstDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);

            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = LastDay;
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\觀光賣場業績及團客散客車數.frx");

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
                            SELECT TA001 AS '日期',TA002 AS '賣場',TMONEYS AS '銷售金額',GROUPMONEYS AS '團客',VISITORMONEYS AS '散客',CARNUM AS '來車數'
                            ,(CASE WHEN GROUPMONEYS>0 AND CARNUM>0 THEN CONVERT(decimal(16,0),GROUPMONEYS/CARNUM) ELSE  0 END ) AS '平均每車金額'
                            ,(SELECT ISNULL(SUM(TB031) ,0)
                            FROM [TK].dbo.POSTB
                            WHERE TB010 LIKE '406%'
                            AND TB002=TA002
                            AND TB001=TA001) AS '霜淇淋金額'
                            ,(TMONEYS-(SELECT ISNULL(SUM(TB031) ,0)
                            FROM [TK].dbo.POSTB
                            WHERE TB010 LIKE '406%'
                            AND TB002=TA002
                            AND TB001=TA001)) AS '銷售金額扣霜淇淋'
                            FROM 
                            (
                            SELECT 
                            TA001,TA002 ,SUM(TA026) AS 'TMONEYS'
                            ,(SELECT ISNULL(SUM(TA026),0) FROM  [TK].dbo.POSTA TA WHERE TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' OR TA009 LIKE '68%' OR TA009 LIKE '69%' )) AS 'GROUPMONEYS'
                            ,(SUM(TA026)-(SELECT ISNULL(SUM(TA026),0) FROM  [TK].dbo.POSTA TA WHERE TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' OR TA009 LIKE '68%' OR TA009 LIKE '69%' ) ) ) AS 'VISITORMONEYS'
                            ,CASE WHEN TA002 IN ('106701') THEN (SELECT ISNULL(SUM(CARNUM),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001) ELSE 0 END  AS 'CARNUM'
                            FROM [TK].dbo.POSTA
                            WHERE TA002 IN ('106701')
                            AND TA001>='{0}' AND TA001<='{1}'
                            GROUP BY TA002,TA001
                            ) AS TEMP
                            ORDER BY TA001,TA002
 

                            ", SDATES, EDATES);

            return SB;

        }

        public void Search(string SDAYS, string EDAYS)
        {
            StringBuilder sbSql = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                sbSqlQuery.Clear();
                              
                sbSql.AppendFormat(@"  
                                    SELECT TB001,'麵包類' AS  '麵包類' ,'全部' AS '分類',CONVERT(INT,SUM(TB031)) AS '未稅金額'
                                    ,(SELECT ISNULL(CONVERT(INT,SUM(TB031)),0)
                                    FROM [TK].dbo.POSTA TA,[TK].dbo.POSTB TB
                                    WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006
                                    AND TB.TB002 IN ('106701')
                                    AND (TB.TB010 LIKE '408%' OR TB.TB010 LIKE '409%')
                                    AND ( TA.TA008 LIKE '68%' OR TA.TA008 LIKE '69%')
                                    AND TB.TB001=POSTB.TB001) AS '團客'
                                    ,(SELECT ISNULL(CONVERT(INT,SUM(TB031)),0)
                                    FROM [TK].dbo.POSTA TA,[TK].dbo.POSTB TB
                                    WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006
                                    AND TB.TB002 IN ('106701')
                                    AND (TB.TB010 LIKE '408%' OR TB.TB010 LIKE '409%')
                                    AND  TA.TA008 NOT LIKE '68%' 
                                    AND  TA.TA008 NOT LIKE '69%'
                                    AND  TB.TB001=POSTB.TB001) AS '散客'
                                    FROM [TK].dbo.POSTA,[TK].dbo.POSTB
                                    WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003 AND TA006=TB006
                                    AND TB001>='{0}' AND TB001<='{1}'
                                    AND TB002 IN ('106701')
                                    AND (TB010 LIKE '408%' OR TB010 LIKE '409%')
                                    GROUP BY TB001
                                    ORDER BY TB001
                                    ", SDAYS,EDAYS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        // 將 "未稅金額" 欄位格式化為金錢
                        dataGridView1.Columns["未稅金額"].DefaultCellStyle.Format = "N0";
                        dataGridView1.Columns["團客"].DefaultCellStyle.Format = "N0";
                        dataGridView1.Columns["散客"].DefaultCellStyle.Format = "N0";

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SETFASTREPORT_DAILY(string SDATES)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();
            StringBuilder SQL4 = new StringBuilder();
            StringBuilder SQL5 = new StringBuilder();
            StringBuilder SQL6 = new StringBuilder();

            SQL1 = SETSQL_DAILY1(SDATES);
            SQL2 = SETSQL_DAILY2(SDATES);
            SQL3 = SETSQL_DAILY3(SDATES);
            SQL4 = SETSQL_DAILY4(SDATES);
            SQL5 = SETSQL_DAILY5(SDATES);
            SQL6 = SETSQL_DAILY6(SDATES);

            Report report1 = new Report();
            report1.Load(@"REPORT\觀光日報表.frx");

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
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();
            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();
            TableDataSource table3 = report1.GetDataSource("Table3") as TableDataSource;
            table3.SelectCommand = SQL4.ToString();
            TableDataSource table4 = report1.GetDataSource("Table4") as TableDataSource;
            table4.SelectCommand = SQL5.ToString();
            TableDataSource table5 = report1.GetDataSource("Table5") as TableDataSource;
            table5.SelectCommand = SQL6.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL_DAILY1(string SDATES)
        {
            StringBuilder SB = new StringBuilder();

             
            SB.AppendFormat(@"                              
                            SELECT 
                            TA001 AS '日期'
                            ,TA002 AS '門市'
                            ,SUM(TA026) AS '方城市合計'
                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TA002 AND TB010 LIKE '406%'),0) AS '霜淇淋業績'
                            ,ISNULL((SUM(TA026)-(SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TB002 AND TB010 LIKE '406%')),0) AS '方塊酥業績'
                            ,ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001=POSTA.TA001 AND TA1.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' )),0) AS '團客業績'
                            ,ISNULL((SUM(TA026)-(ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001=POSTA.TA001 AND TA1.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' )),0)))-(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TA002 AND TB010 LIKE '406%'),0) AS '散客業績'
                            ,(SELECT COUNT([ID]) FROM  [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=TA001 AND STATUS='完成接團') AS '車數'
                            ,CASE WHEN SUM(TA026)>0 AND (SELECT COUNT([ID]) FROM  [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=TA001 AND STATUS='完成接團')>0 THEN CONVERT(INT,ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001=POSTA.TA001 AND TA1.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' )),0)/(SELECT COUNT([ID]) FROM  [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=TA001 AND STATUS='完成接團')) ELSE 0 END '平均每車金額'
                            ,ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001>=CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112) AND TA1.TA001<=POSTA.TA001 AND TA1.TA002=POSTA.TA002 ),0) AS '目前累計'
                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TA002 AND ( TB010 LIKE '408%' OR  TB010 LIKE '409%' OR  TB010 LIKE '40400610020011%')),0) AS '硯微墨的寄賣'

                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002  IN ('106703') ),0) AS '星球合計'
                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002  IN ('106703') AND TB010 LIKE '598%'),0) AS '星球業績'
                            ,ISNULL(((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002  IN ('106703') )-(SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002  IN ('106703') AND TB010 LIKE '598%')),0) AS '其他業績'
                            ,ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001>=CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112) AND TA1.TA001<=POSTA.TA001 AND TA1.TA002 IN ('106703') ),0) AS '星球樂園目前累計'

                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB TB2,[TK].dbo.POSTA TA2 WHERE TA2.TA001=TB2.TB001 AND  TA2.TA002=TB2.TB002 AND TA2.TA003=TB2.TB003 AND TA2.TA006=TB2.TB006  AND TA2.TA038='4' AND TB2.TB001=POSTA.TA001 AND TB2.TB002=POSTA.TA002 ),0) AS '預購業績'
                            FROM [TK].dbo.POSTA
                            WHERE 1=1
                            AND TA002 IN ('106701')
                            AND TA001='{0}'
                            GROUP BY TA001,TA002 
                            ORDER BY TA001,TA002
 

                            ", SDATES);

            return SB;

        }
        public StringBuilder SETSQL_DAILY2(string SDATES)
        { 
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                            SELECT
                            TA001 AS '日期'
                            ,TA002 AS '門市'
                            , ISNULL(SUM(TA026),0) AS '硯微墨烘焙組合計'
                            ,ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001>=CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112) AND TA1.TA001<=POSTA.TA001 AND TA1.TA002=POSTA.TA002 ),0) AS '目前累計'
                            ,ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001=POSTA.TA001 AND TA1.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' )),0) AS '團客業績'
                            ,ISNULL((SUM(TA026)-(ISNULL((SELECT SUM(TA026) FROM [TK].dbo.POSTA TA1 WHERE TA1.TA001=POSTA.TA001 AND TA1.TA002=POSTA.TA002 AND (TA008 LIKE '68%' OR TA008 LIKE '69%' )),0)))-(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TA002 AND TB010 LIKE '406%'),0) AS '散客業績'

                            FROM [TK].dbo.POSTA
                            WHERE 1=1
                            AND TA002 IN ('106702')
                            AND TA001='{0}'
                            GROUP BY TA001,TA002
                            ORDER BY TA001,TA002
 

                            ", SDATES);

            return SB;

        }
        public StringBuilder SETSQL_DAILY3(string SDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                             
                            
                            WITH 累計值 AS (
                                SELECT 
                                    TA002,
                                    ISNULL(SUM(TA026), 0) AS '目前累計'
                                FROM [TK].dbo.POSTA TA1
                                WHERE TA1.TA001 >= CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112)
                                AND TA1.TA001 <= '{0}'
                                AND TA1.TA002 IN ('106705')
                                GROUP BY TA002
                            )
                            SELECT 
                                '{0}' AS '日期',
                                '106705' AS '門市',
                                ISNULL(本日資料.硯微墨餐飲組合計, 0) AS '硯微墨餐飲組合計',
                                累計值.目前累計 AS '目前累計',
                                ISNULL(本日資料.霜淇淋業績, 0) AS '霜淇淋業績',
                                ISNULL(本日資料.飲品業績, 0) AS '飲品業績',
                                ISNULL(本日資料.其他, 0) AS '其他'
                            FROM 累計值
                            LEFT JOIN (
                                SELECT 
                                    TA001,
                                    TA002,
                                    ISNULL(SUM(TA026), 0) AS '硯微墨餐飲組合計',
                                    ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001 = TA001 AND TB002 = TA002 AND TB010 LIKE '406%'), 0) AS '霜淇淋業績',
                                    ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001 = TA001 AND TB002 = TA002 AND TB010 LIKE '407%'), 0) AS '飲品業績',
                                    (ISNULL(SUM(TA026), 0) 
                                        - ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001 = TA001 AND TB002 = TA002 AND TB010 LIKE '406%'), 0) 
                                        - ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001 = TA001 AND TB002 = TA002 AND TB010 LIKE '407%'), 0)) AS '其他'
                                FROM [TK].dbo.POSTA
                                WHERE TA002 = '106705'
                                AND TA001 = '{0}'
                                GROUP BY TA001, TA002
                            ) AS 本日資料 ON 本日資料.TA002 = 累計值.TA002


                            ", SDATES);

            return SB;

        }

        public StringBuilder SETSQL_DAILY4(string SDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                           WITH 累計值 AS (
                                SELECT 
                                    TA002,
                                    ISNULL(SUM(TA026), 0) AS '目前累計'
                                FROM [TK].dbo.POSTA TA1
                                WHERE TA1.TA001 >= CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112)
                                AND TA1.TA001 <= '{0}'
                                AND TA1.TA002 IN ('106703')
                                GROUP BY TA002
                            )
                            SELECT 
                                '{0}' AS '日期',
                                '106703' AS '門市',
                                ISNULL(本日資料.星球樂園合計, 0) AS '星球樂園合計',
                                ISNULL(本日資料.星球業績, 0) AS '星球業績',
                                ISNULL(本日資料.其他業績, 0) AS '其他業績',
                                累計值.目前累計 AS '目前累計'
                            FROM 累計值
                            LEFT JOIN (
                                SELECT 
                                    TA001,
                                    TA002,
                                    ISNULL(SUM(TA026), 0) AS '星球樂園合計',
                                    (SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TA002 AND TB010 LIKE '598%') AS '星球業績',
                                    (SUM(TA026) - (SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB001=TA001 AND TB002=TA002 AND TB010 LIKE '598%')) AS '其他業績'
                                FROM [TK].dbo.POSTA
                                WHERE TA002 = '106703'
                                AND TA001 = '{0}'
                                GROUP BY TA001, TA002
                            ) AS 本日資料 ON 本日資料.TA002 = 累計值.TA002


                            ", SDATES);

            return SB;

        }
        //67000016 VIP優惠
        //（9折）
        public StringBuilder SETSQL_DAILY5(string SDATES)
        {
            string TA009 = "67000016";
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                               WITH 累計值 AS (
                            SELECT 
                                ISNULL(SUM(TA026), 0) AS '目前累計'
                            FROM [TK].dbo.POSTA TA1
                            WHERE TA1.TA001 >= CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112)
                            AND TA1.TA001 <= '{0}'
                            AND TA1.TA009 = '{1}'
                            )

                            SELECT 
                                '20240904' AS '日期',
                                ISNULL(本日優惠.TA026, 0) AS '67000016(9折)VIP優惠',
                                累計值.目前累計 AS '目前累計'
                            FROM 累計值
                            LEFT JOIN (
                                SELECT SUM(TA026) AS TA026
                                FROM [TK].dbo.POSTA
                                WHERE TA002 LIKE '1067%'
                                AND TA009 = '{1}'
                                AND TA001 = '{0}'
                            ) AS 本日優惠 ON 1 = 1
                                           

                            ", SDATES, TA009);

            return SB;

        }
        //67000017VVIP優惠
        //（85折)
        public StringBuilder SETSQL_DAILY6(string SDATES)
        {
            string TA009 = "67000017";
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  

                            WITH 累計值 AS (
                            SELECT 
                                ISNULL(SUM(TA026), 0) AS '目前累計'
                            FROM [TK].dbo.POSTA TA1
                            WHERE TA1.TA001 >= CONVERT(varchar(8), DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0), 112)
                            AND TA1.TA001 <= '{0}'
                            AND TA1.TA009 = '{1}'
                            )

                            SELECT 
                                '20240904' AS '日期',
                                ISNULL(本日優惠.TA026, 0) AS '67000017(85折)VVIP優惠',
                                累計值.目前累計 AS '目前累計'
                            FROM 累計值
                            LEFT JOIN (
                                SELECT SUM(TA026) AS TA026
                                FROM [TK].dbo.POSTA
                                WHERE TA002 LIKE '1067%'
                                AND TA009 = '{1}'
                                AND TA001 = '{0}'
                            ) AS 本日優惠 ON 1 = 1

                  

                            ", SDATES ,TA009);

            return SB;

        }


        #endregion

        #region BUTTON

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_DAILY(dateTimePicker5.Value.ToString("yyyyMMdd"));
        }
        #endregion

    }
}
