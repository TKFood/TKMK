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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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

        public void ADDCARKIND(string ID,string NAME)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
            STATUSCARKIND = "EDIT";
            SETTEXT2();
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

            SEARCHCARKIND();
            SETTEXT3();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETTEXT3();
        }

        #endregion


    }
}
