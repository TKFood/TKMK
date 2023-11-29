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
            report1.Load(@"REPORT\觀光賣場的團客散客車數.frx");

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
        #endregion


    }
}
