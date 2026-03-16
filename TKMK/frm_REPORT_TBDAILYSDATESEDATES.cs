using FastReport;
using FastReport.Data;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.Extractor;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TKITDLL;

namespace TKMK
{
    public partial class frm_REPORT_TBDAILYSDATESEDATES : Form
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

        public frm_REPORT_TBDAILYSDATESEDATES()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void ADD_TBDAILYSDATESEDATES(string SDATES,string EDATES)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (sqlConn = new SqlConnection(sqlsb.ConnectionString))
            using (cmd = new SqlCommand())
            {
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                DELETE [TKMK].[dbo].[TBDAILYSDATESEDATES]
                                WHERE [SDATES]=@SDATES AND [EDATES]=@EDATES

                                INSERT INTO [TKMK].[dbo].[TBDAILYSDATESEDATES]
                                (
                                [SDATES]
                                ,[EDATES]
                                ,[MB001]
                                ,[MB002]
                                ,[期初庫存]
                                ,[期末庫存]
                                ,[本期銷售]
                                ,[本期入庫]
                                ,[本期領用]
                                ,[本期轉撥入]
                                ,[本期轉撥出]
                                )

                                SELECT 
                                @SDATES AS SDATES
                                ,@EDATES AS EDATES
                                ,MB001
                                ,MB002
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)
                                FROM [TK].dbo.INVLA   WITH(NOLOCK)
                                WHERE  LA001=TEMP.MB001 and LA004<@SDATES 
                                AND (LA009 ='21002' )),0) AS '期初庫存'
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)
                                FROM [TK].dbo.INVLA 
                                WHERE  LA001=TEMP.MB001 and LA004<=@EDATES 
                                AND (LA009 ='21002' )),0) AS '期末庫存'
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)*-1
                                FROM [TK].dbo.INVLA   WITH(NOLOCK)
                                WHERE  LA001=TEMP.MB001 
                                AND LA014 IN ('2')
                                AND LA004>=@SDATES 
                                AND LA004<=@EDATES 
                                AND (LA009 ='21002' )),0) AS '本期銷售'
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)
                                FROM [TK].dbo.INVLA   WITH(NOLOCK)
                                WHERE  LA001=TEMP.MB001 
                                AND LA014 IN ('1')
                                AND LA004>=@SDATES 
                                AND LA004<=@EDATES 
                                AND (LA009 ='21002' )),0) AS '本期入庫'
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)*-1
                                FROM [TK].dbo.INVLA   WITH(NOLOCK)
                                WHERE  LA001=TEMP.MB001 
                                AND LA014 IN ('3')
                                AND LA004>=@SDATES 
                                AND LA004<=@EDATES 
                                AND (LA009 ='21002' )),0) AS '本期領用'
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)
                                FROM [TK].dbo.INVLA   WITH(NOLOCK)
                                WHERE  LA001=TEMP.MB001 
                                AND LA014 IN ('4')
                                AND LA005 IN (1)
                                AND LA004>=@SDATES 
                                AND LA004<=@SDATES 
                                AND (LA009 ='21002' )),0) AS '本期轉撥入'
                                ,ISNULL((
                                SELECT SUM(LA011*LA005)*-1
                                FROM [TK].dbo.INVLA   WITH(NOLOCK)
                                WHERE  LA001=TEMP.MB001 
                                AND LA014 IN ('4')
                                AND LA005 IN (-1)
                                AND LA004>=@SDATES 
                                AND LA004<=@EDATES 
                                AND (LA009 ='21002' )),0) AS '本期轉撥出'

                                FROM 
                                (
	                                SELECT 
	                                DISTINCT A.MA002 AS MA002, A.MA003 AS MA003, MB001,INV.MC002 AS LA009, A.MA004 AS MA004, ISNULL(B.MA003,'') AS ACTMA003, MB002, MB003, MB004, MB072, CMS.MC002 AS CMSMC002, MB090,ISNULL(MD003,0) AS MD003,ISNULL(MD004,0) AS MD004 
	                                FROM  [TK].dbo.INVMB AS INVMB  WITH(NOLOCK)
	                                INNER JOIN [TK].dbo.INVMC AS INV  WITH(NOLOCK) ON INV.MC001=MB001 
	                                LEFT JOIN  [TK].dbo.CMSMC AS CMS  WITH(NOLOCK) ON INV.MC002=CMS.MC001 
	                                LEFT JOIN  [TK].dbo.INVMD AS INVMD  WITH(NOLOCK) ON MB001=MD001 AND MB072=MD002 
	                                INNER JOIN  [TK].dbo.INVLA AS INVLA  WITH(NOLOCK) ON LA001=MB001 AND LA009=CMS.MC001
	                                LEFT JOIN  [TK].dbo.INVMA AS A  WITH(NOLOCK) ON A.MA001='1' AND A.MA002=MB005 
									LEFT JOIN  [TK].dbo.ACTMC AS ACTMC  WITH(NOLOCK) ON 1=1 
									LEFT JOIN  [TK].dbo.ACTMA AS B  WITH(NOLOCK) ON B.MA001=A.MA004 AND B.MA050=ACTMC.MC039
									Where  (LA004 Between @SDATES and @EDATES)  
									AND (INV.MC002 IN (N'21002'))
	                                AND CMS.MC004='1'  
	                                AND ISNULL(A.MA001,'')<>'' 
	                                AND ISNULL(A.MA002,'')<>'' 
                                    AND (MB001 LIKE '4%' OR MB001 LIKE '5%')
                                ) AS TEMP
                                ORDER BY  MB001,MB002


                                UPDATE [TKMK].[dbo].[TBDAILYSDATESEDATES]
                                SET [其他]=[期末庫存]+[本期銷售]-[本期入庫]+[本期領用]-[本期轉撥入]+[本期轉撥出]-[期初庫存]
                                WHERE  [SDATES]=@SDATES AND [EDATES]=@EDATES


                                UPDATE [TKMK].[dbo].[TBDAILYSDATESEDATES]
                                SET [販售率] = CAST(
                                                CASE 
                                                    WHEN ([本期入庫] + [期初庫存] + [本期轉撥入]+[其他]) = 0 THEN 0 
                                                    ELSE ([本期銷售] * 100.0) / ([本期入庫] + [期初庫存] + [本期轉撥入]+[其他])
                                                END 
                                                AS DECIMAL(18, 2)
                                              )
                                WHERE [SDATES] = @SDATES 
                                  AND [EDATES] = @EDATES
                                ");
                cmd.CommandText = sbSql.ToString();
                cmd.Connection = sqlConn;
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@SDATES", SDATES);
                cmd.Parameters.AddWithValue("@EDATES", EDATES);
                try
                {
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();
                    tran.Commit();
                }
                catch (Exception ex)
                {
                    if (tran != null)
                        tran.Rollback();
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\\硯微墨商品販售率.frx");

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
                            [SDATES] AS '日期起'
                            ,[EDATES] AS '日期迄'
                            ,[MB001] AS '品號' 
                            ,[MB002] AS '品名'
                            ,[期初庫存]
                            ,[期末庫存]
                            ,[本期銷售]
                            ,[本期入庫]
                            ,[本期領用]
                            ,[本期轉撥入]
                            ,[本期轉撥出]
                            ,[其他]
                            ,[販售率]    

                            FROM [TKMK].[dbo].[TBDAILYSDATESEDATES]
                            WHERE [SDATES]='{0}' AND [EDATES]='{1}'
                            ORDER BY [SDATES], [EDATES],[MB001]


                            ", SDATES, EDATES);

            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker2.Value.ToString("yyyyMMdd");

            ADD_TBDAILYSDATESEDATES(SDATES, EDATES);
            SETFASTREPORT(SDATES, EDATES);
        }
        #endregion
    }
}
