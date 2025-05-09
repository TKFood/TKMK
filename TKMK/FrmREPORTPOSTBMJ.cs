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
    public partial class FrmREPORTPOSTBMJ : Form
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

        public FrmREPORTPOSTBMJ()
        {
            InitializeComponent();
        }

      

        private void FrmREPORTPOSTBMJ_Load(object sender, EventArgs e)
        {
            //SEARCH();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDATES, string EDATES,string YEARS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES, YEARS);
            Report report1 = new Report();
            report1.Load(@"REPORT\觀光查活動組合.frx");

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

        public StringBuilder SETSQL(string SDATES, string EDATES,string YEARS)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 
                            TA002 AS '門市'
                            ,TA001 AS '日期'
                            ,TA003 AS '機台'
                            ,TA014 AS '發票'
                            ,TB019 AS '銷售數量'
                            ,TB025 AS '促銷折扣金額'
                            ,ML004005 AS '折扣件數'
                            FROM 
                            (
                            SELECT 
                            TA002,TA001,TA003,TA014,SUM(TB019) 'TB019',SUM(TB025) 'TB025'
	                            ,(SELECT SUM(ML004+ML005)
	                            FROM [TK].dbo.POSML
	                            WHERE ML003='420250101016') AS 'ML004005'
                            FROM [TK].dbo.POSTA WITH(NOLOCK),[TK].dbo.POSTB WITH(NOLOCK)
                            WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003 AND TA006=TB006
                            AND TA002 IN (SELECT  [TA002]  FROM [TKMK].[dbo].[TB_TA002])
                            AND ISNULL(TA014,'')<>''
                            AND TB010 IN 
                            (
	                            SELECT MJ004
	                            FROM [TK].dbo.POSMJ
	                            WHERE MJ003 IN (SELECT  [MJ003]  FROM [TKMK].[dbo].[TB_MJ003] WHERE YEARS='{2}')
                            )
                            GROUP BY TA002,TA001,TA003,TA014
                            ) AS TEMP 
                            WHERE (TB019 % ML004005 <> 0 OR ML004005 = 0)
                            AND TB019>=5
                            AND TA001>='{0}' AND TA001<='{1}'
                            ORDER BY TA002,TA001,TA003,TA014
 

                            ", SDATES, EDATES, YEARS);

            return SB;

        }

        public void SEARCH()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT 
                                    [MJ003] AS '活動代號'
                                    ,[NAMES] AS '名稱'
                                    ,[YEARS] AS '年度'
                                    FROM [TKMK].[dbo].[TB_MJ003]
                                    ORDER BY [YEARS]

                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();                      
                    }

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


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SETNULL();

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    string MJ003 = row.Cells["活動代號"].Value.ToString();

                    textBox4.Text = MJ003;
                }
            }
        }

        public DataTable FIND_TB_MJ003(string YEARS)
        {
            DataTable DT = new DataTable();

            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                   SELECT 
                                    [MJ003] AS '活動代號'
                                    ,[NAMES] AS '名稱'
                                    ,[YEARS] AS '年度'
                                    FROM [TKMK].[dbo].[TB_MJ003]
                                    WHERE [YEARS]='{0}'
                                    ORDER BY [YEARS]

                                    ", YEARS);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void ADD_TB_MJ003(
            string MJ003
            ,string NAMES
            ,string YEARS
            )
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
                                    INSERT INTO [TKMK].[dbo].[TB_MJ003]
                                    (
                                    [MJ003]
                                    ,[NAMES]
                                    ,[YEARS]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    )
                                    ", MJ003
                                    , NAMES
                                    , YEARS
                                    );

                sbSql.AppendFormat(@" ");

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


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("失敗 " + ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELETE_TB_MJ003(
           string MJ003        
           )
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
                                    DELETE [TKMK].[dbo].[TB_MJ003]
                                    WHERE  [MJ003]='{0}'                                   
                                    ", MJ003
                                   
                                    );

                sbSql.AppendFormat(@" ");

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


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("失敗 " + ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void SETNULL()
        {
            textBox4.Text = "";
        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker1.Value.ToString("yyyy"));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string YEARS = textBox2.Text.Trim();
            string MJ003 = textBox1.Text.Trim();
            string NAMES = textBox3.Text.Trim();
            DataTable DT = FIND_TB_MJ003(YEARS);

            if(!string.IsNullOrEmpty(YEARS))
            {
                if (DT != null && DT.Rows.Count >= 1)
                {
                    MessageBox.Show("同年度不可有2個活動檢查設定");
                }
                else
                {
                    ADD_TB_MJ003(
                        MJ003
                        , NAMES
                        , YEARS
                        );
                    SEARCH();
                    MessageBox.Show("完成");
                }
            }
            else
            {
                MessageBox.Show("活動代號、名稱、年度 不得空白");
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string MJ003 = textBox4.Text.Trim();
                if(!string.IsNullOrEmpty(MJ003))
                {
                    DELETE_TB_MJ003(MJ003);
                    SEARCH();
                    MessageBox.Show("完成");
                }                

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        #endregion

     
    }
}
