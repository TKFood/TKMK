using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKMK
{
    public partial class frmGROUPSET : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        int result;

        string STATUSCARKIND;
        string IDCARKIND;
        string STATUSGROUPBASE;
        string IDGROUPBASE;
        string STATUSGROUPPCT;
        string IDGROUPPCT;
        string STATUSGROUPPRODUCT;
        string IDGROUPPRODUCT;
        string STATUSGROUPEXCHANGEMONEYS;
        string IDGROUPEXCHANGEMONEYS;


        public frmGROUPSET()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHCARKIND()
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

                sbSql.AppendFormat(@"  SELECT [ID] AS '代號',[NAME] AS '名稱'  FROM [TKMK].[dbo].[CARKIND] ORDER BY [ID]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public void SEARCHGROUPBASE()
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

                sbSql.AppendFormat(@"  SELECT [ID] AS '代號',[NAME] AS '名稱',[BASEMONEYS] AS '茶水費'");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPBASE]");
                sbSql.AppendFormat(@"  ORDER BY [ID]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds1.Tables["ds1"];
                        dataGridView2.AutoResizeColumns();


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

        public void SEARCHGROUPPCT()
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

                sbSql.AppendFormat(@"  SELECT [MID] AS '代號', [ID] AS '車種',[PCTMONEYS] AS '消費金額門檻',[NAME] AS '名稱',[PCT] AS '抽佣比例'");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPPCT]");
                sbSql.AppendFormat(@"  ORDER BY [ID],[NAME],[PCTMONEYS]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds1.Tables["ds1"];
                        dataGridView3.AutoResizeColumns();


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

        public void SEARCHGROUPPRODUCT()
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
                                    SELECT  [ID] AS '代號',[NAME] AS '名稱',[NUM] AS '組數',[MONEYS] AS '金額',[SPLITCAL] AS '是否分拆計算',[VALID] AS '是否使用'
                                    FROM [TKMK].[dbo].[GROUPPRODUCT]
                                    ORDER BY [ID]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds1.Tables["ds1"];
                        dataGridView4.AutoResizeColumns();


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

        public void SEARCHGROUPEXCHANGEMONEYS()
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

                sbSql.AppendFormat(@"  SELECT [ID] AS '代號',[EXCHANGEMONEYS] AS '兌換券金額'");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPEXCHANGEMONEYS]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView5.DataSource = ds1.Tables["ds1"];
                        dataGridView5.AutoResizeColumns();


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
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    IDCARKIND = row.Cells["代號"].Value.ToString();
                    textBox11.Text = row.Cells["代號"].Value.ToString();
                    textBox12.Text = row.Cells["名稱"].Value.ToString();
                }
                else
                {
                    IDCARKIND = null;
                    textBox11.Text = null;
                    textBox12.Text = null;
                }
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    IDGROUPBASE = row.Cells["代號"].Value.ToString();
                    textBox21.Text = row.Cells["代號"].Value.ToString();
                    textBox22.Text = row.Cells["名稱"].Value.ToString();
                    textBox23.Text = row.Cells["茶水費"].Value.ToString();
                }
                else
                {
                    IDGROUPBASE = null;
                    textBox21.Text = null;
                    textBox22.Text = null;
                    textBox23.Text = null;
                }
            }
        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    IDGROUPPCT = row.Cells["代號"].Value.ToString();
                    textBox31.Text = row.Cells["代號"].Value.ToString();
                    textBox32.Text = row.Cells["車種"].Value.ToString();
                    textBox33.Text = row.Cells["消費金額門檻"].Value.ToString();
                    textBox34.Text = row.Cells["名稱"].Value.ToString();
                    textBox35.Text = row.Cells["抽佣比例"].Value.ToString();
                }
                else
                {
                    IDGROUPPCT = null;
                    textBox31.Text = null;
                    textBox32.Text = null;
                    textBox33.Text = null;
                    textBox34.Text = null;
                    textBox35.Text = null;
                }
            }
        }
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    IDGROUPPRODUCT = row.Cells["代號"].Value.ToString();
                    textBox41.Text = row.Cells["代號"].Value.ToString();
                    textBox42.Text = row.Cells["名稱"].Value.ToString();
                    textBox43.Text = row.Cells["組數"].Value.ToString();
                    textBox44.Text = row.Cells["金額"].Value.ToString();                   
                    textBox46.Text = row.Cells["是否分拆計算"].Value.ToString();
                    textBox47.Text = row.Cells["是否使用"].Value.ToString();
                }
                else
                {
                    IDGROUPPRODUCT = null;
                    textBox41.Text = null;
                    textBox42.Text = null;
                    textBox43.Text = null;
                    textBox44.Text = null;
                    
                    textBox46.Text = null;
                    textBox47.Text = null;
                }
            }

        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    IDGROUPEXCHANGEMONEYS = row.Cells["代號"].Value.ToString();
                    textBox51.Text = row.Cells["代號"].Value.ToString();
                    textBox52.Text = row.Cells["兌換券金額"].Value.ToString();

                }
                else
                {
                    IDGROUPEXCHANGEMONEYS = null;
                    textBox51.Text = null;
                    textBox52.Text = null;
                   
                }
            }
        }

        public void ADDCARKIND(string ID,string NAME)
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

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[CARKIND]");
                sbSql.AppendFormat(" ([ID],[NAME])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}')", ID, NAME);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATECARKIND(string IDCARKIND, string ID, string NAME)
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

                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[CARKIND] ");
                sbSql.AppendFormat(" SET [ID]='{0}',[NAME]='{1}'",ID,NAME);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", IDCARKIND);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDGROUPBASE(string ID, string NAME,string BASEMONEYS)
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

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[GROUPBASE]");
                sbSql.AppendFormat(" ([ID],[NAME],[BASEMONEYS])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}','{2}')",ID,NAME,BASEMONEYS);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEGROUPBASE(string IDGROUPBASE, string ID, string NAME, string BASEMONEYS)
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


                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPBASE]");
                sbSql.AppendFormat(" SET [ID]='{0}',[NAME]='{1}',[BASEMONEYS]='{2}'", ID, NAME, BASEMONEYS);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", IDGROUPBASE);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDGROUPPCT(string MID, string ID, string PCTMONEYS, string NAME, string PCT)
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

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[GROUPPCT]");
                sbSql.AppendFormat(" ([MID],[ID],[PCTMONEYS],[NAME],[PCT])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}')", MID, ID,PCTMONEYS,NAME,PCT);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEGROUPPCT(string IDGROUPPCT, string MID, string ID, string PCTMONEYS, string NAME, string PCT)
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


                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPPCT]");
                sbSql.AppendFormat(" SET [MID]='{0}',[ID]='{1}',[PCTMONEYS]='{2}',[NAME]='{3}',[PCT]='{4}'",MID, ID, PCTMONEYS, NAME, PCT);
                sbSql.AppendFormat(" WHERE [MID]='{0}'", IDGROUPPCT);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDGROUPPRODUCT(string ID, string NAME, string NUM, string MONEYS, string SPLITCAL,string VALID)
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


                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[GROUPPRODUCT]");
                sbSql.AppendFormat(" ([ID],[NAME],[NUM],[MONEYS],[SPLITCAL],[VALID])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}')", ID, NAME, NUM, MONEYS, SPLITCAL, VALID);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEGROUPPRODUCT(string IDGROUPPRODUCT, string ID, string NAME, string NUM, string MONEYS, string SPLITCAL,string VALID)
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

                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPPRODUCT]");
                sbSql.AppendFormat(" SET [ID]='{0}',[NAME]='{1}',[NUM]='{2}',[MONEYS]='{3}',[SPLITCAL]='{4}',[VALID]='{5}'", ID, NAME, NUM, MONEYS, SPLITCAL, VALID);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", IDGROUPPRODUCT);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDGROUPEXCHANGEMONEYS(string ID, string EXCHANGEMONEYS)
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

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[GROUPEXCHANGEMONEYS]");
                sbSql.AppendFormat(" ([ID],[EXCHANGEMONEYS])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}')", ID, EXCHANGEMONEYS);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEGROUPEXCHANGEMONEYS(string IDGROUPEXCHANGEMONEYS, string ID, string EXCHANGEMONEYS)
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

                sbSql.AppendFormat(" UPDATE  [TKMK].[dbo].[GROUPEXCHANGEMONEYS]");
                sbSql.AppendFormat(" SET [ID]='{0}',[EXCHANGEMONEYS]='{1}'", ID, EXCHANGEMONEYS);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", IDGROUPEXCHANGEMONEYS);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SETTEXT1()
        {
            textBox11.ReadOnly = false;
            textBox12.ReadOnly = false;

            textBox11.Text = null;
            textBox12.Text = null;
        }

        public void SETTEXT2()
        {
            textBox11.ReadOnly = false;
            textBox12.ReadOnly = false;


        }

        public void SETTEXT3()
        {
            textBox11.ReadOnly = true;
            textBox12.ReadOnly = true;


        }

        public void SETTEXT4()
        {
            textBox21.ReadOnly = false;
            textBox22.ReadOnly = false;
            textBox23.ReadOnly = false;

            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
        }

        public void SETTEXT5()
        {
            textBox21.ReadOnly = false;
            textBox22.ReadOnly = false;
            textBox23.ReadOnly = false;

        }

        public void SETTEXT6()
        {
            textBox21.ReadOnly = true;
            textBox22.ReadOnly = true;
            textBox23.ReadOnly = true;

        }

        public void SETTEXT7()
        {
            textBox31.ReadOnly = false;
            textBox32.ReadOnly = false;
            textBox33.ReadOnly = false;
            textBox34.ReadOnly = false;
            textBox35.ReadOnly = false;

            textBox31.Text = null;
            textBox32.Text = null;
            textBox33.Text = null;
            textBox34.Text = null;
            textBox35.Text = null;
        }

        public void SETTEXT8()
        {
            textBox31.ReadOnly = false;
            textBox32.ReadOnly = false;
            textBox33.ReadOnly = false;
            textBox34.ReadOnly = false;
            textBox35.ReadOnly = false;
        }

        public void SETTEXT9()
        {
            textBox31.ReadOnly = true;
            textBox32.ReadOnly = true;
            textBox33.ReadOnly = true;
            textBox34.ReadOnly = true;
            textBox35.ReadOnly = true;

        }

        public void SETTEXT10()
        {
            textBox41.ReadOnly = false;
            textBox42.ReadOnly = false;
            textBox43.ReadOnly = false;
            textBox44.ReadOnly = false;
            
            textBox46.ReadOnly = false;
            textBox47.ReadOnly = false;

            textBox41.Text = null;
            textBox42.Text = null;
            textBox43.Text = null;
            textBox44.Text = null;
            
            textBox46.Text = null;
            textBox47.Text = null;
        }

        public void SETTEXT11()
        {           
            textBox41.ReadOnly = false;
            textBox42.ReadOnly = false;
            textBox43.ReadOnly = false;
            textBox44.ReadOnly = false;
           
            textBox46.ReadOnly = false;
            textBox47.ReadOnly = false;
        }

        public void SETTEXT12()
        {
            textBox41.ReadOnly = true;
            textBox42.ReadOnly = true;
            textBox43.ReadOnly = true;
            textBox44.ReadOnly = true;
           
            textBox46.ReadOnly = true;
            textBox47.ReadOnly = true;
        }


        public void SETTEXT13()
        {
            textBox51.ReadOnly = false;
            textBox52.ReadOnly = false;

            textBox51.Text = null;
            textBox52.Text = null;
            
        }

        public void SETTEXT14()
        {
            textBox51.ReadOnly = false;
            textBox52.ReadOnly = false;

        }

        public void SETTEXT15()
        {
            textBox51.ReadOnly = true;
            textBox52.ReadOnly = true;

        }
        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHCARKIND();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            STATUSCARKIND = "ADD";
            SETTEXT1();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if(STATUSCARKIND.Equals("ADD"))
            {
                ADDCARKIND(textBox11.Text.Trim(), textBox12.Text.Trim());
            }
            else if (STATUSCARKIND.Equals("EDIT"))
            {
                UPDATECARKIND(IDCARKIND,textBox11.Text.Trim(), textBox12.Text.Trim());
            }

            STATUSCARKIND = null;
            SEARCHCARKIND();
            SETTEXT3();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETTEXT3();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHGROUPBASE();
        }
        private void button10_Click(object sender, EventArgs e)
        {
            STATUSGROUPBASE = "ADD";
            SETTEXT4();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            STATUSGROUPBASE = "EDIT";
            SETTEXT5();
        }
        private void button8_Click(object sender, EventArgs e)
        {

            if (STATUSGROUPBASE.Equals("ADD"))
            {
                ADDGROUPBASE(textBox21.Text.Trim(), textBox22.Text.Trim(), textBox23.Text.Trim());
            }
            else if (STATUSGROUPBASE.Equals("EDIT"))
            {
                UPDATEGROUPBASE(IDGROUPBASE, textBox21.Text.Trim(), textBox22.Text.Trim(), textBox23.Text.Trim());
            }

            STATUSGROUPBASE = null;
            SEARCHGROUPBASE();
            SETTEXT6();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SETTEXT6();
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHGROUPPCT();
        }


        private void button15_Click(object sender, EventArgs e)
        {
            STATUSGROUPPCT = "ADD";
            SETTEXT7();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            STATUSGROUPPCT = "EDIT";
            SETTEXT8();
        }

        private void button13_Click(object sender, EventArgs e)
        {

            if (STATUSGROUPPCT.Equals("ADD"))
            {
                ADDGROUPPCT(textBox31.Text.Trim(), textBox32.Text.Trim(), textBox33.Text.Trim(), textBox34.Text.Trim(), textBox35.Text.Trim());
            }
            else if (STATUSGROUPPCT.Equals("EDIT"))
            {
                UPDATEGROUPPCT(IDGROUPPCT, textBox31.Text.Trim(), textBox32.Text.Trim(), textBox33.Text.Trim(), textBox34.Text.Trim(), textBox35.Text.Trim());
            }

            STATUSGROUPPCT = null;
            SEARCHGROUPPCT();
            SETTEXT9();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SETTEXT9();
        }
        private void button16_Click(object sender, EventArgs e)
        {
            SEARCHGROUPPRODUCT();
        }
        private void button20_Click(object sender, EventArgs e)
        {
            STATUSGROUPPRODUCT = "ADD";
            SETTEXT10();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            STATUSGROUPPRODUCT = "EDIT";
            SETTEXT11();
        }

        private void button18_Click(object sender, EventArgs e)
        {

            if (STATUSGROUPPRODUCT.Equals("ADD"))
            {
                ADDGROUPPRODUCT(textBox41.Text.Trim(), textBox42.Text.Trim(), textBox43.Text.Trim(), textBox44.Text.Trim(),  textBox46.Text.Trim(), textBox47.Text.Trim());
            }
            else if (STATUSGROUPPRODUCT.Equals("EDIT"))
            {
                UPDATEGROUPPRODUCT(IDGROUPPRODUCT, textBox41.Text.Trim(), textBox42.Text.Trim(), textBox43.Text.Trim(), textBox44.Text.Trim(),  textBox46.Text.Trim(), textBox47.Text.Trim());
            }

            STATUSGROUPPRODUCT = null;
            SEARCHGROUPPRODUCT();
            SETTEXT12();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            SETTEXT12();
        }


        private void button21_Click(object sender, EventArgs e)
        {
            SEARCHGROUPEXCHANGEMONEYS();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            STATUSGROUPEXCHANGEMONEYS = "ADD";
            SETTEXT13();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            STATUSGROUPEXCHANGEMONEYS = "EDIT";
            SETTEXT14();
        }

        private void button23_Click(object sender, EventArgs e)
        {

            if (STATUSGROUPEXCHANGEMONEYS.Equals("ADD"))
            {
                ADDGROUPEXCHANGEMONEYS(textBox51.Text.Trim(), textBox52.Text.Trim());
            }
            else if (STATUSGROUPEXCHANGEMONEYS.Equals("EDIT"))
            {
                UPDATEGROUPEXCHANGEMONEYS(IDGROUPEXCHANGEMONEYS, textBox51.Text.Trim(), textBox52.Text.Trim());
            }

            STATUSGROUPEXCHANGEMONEYS = null;
            SEARCHGROUPEXCHANGEMONEYS();
            SETTEXT15();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            SETTEXT15();
        }

        #endregion

       
    }
}
