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
    public partial class frmMKDRINKRECORD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;
        string BUYNO;
        string OLDBUYNO;
        string CHECKYN = "N";

        public frmMKDRINKRECORD()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();

        }

        #region FUNCTION

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ME001";
            comboBox1.DisplayMember = "ME002";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [DRINKNAME] FROM [TKMK].[dbo].[DRINKNAME] WHERE [USED]='Y' ORDER BY [ID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("DRINKNAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "DRINKNAME";
            comboBox2.DisplayMember = "DRINKNAME";
            sqlConn.Close();


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            if(!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString())&&!comboBox1.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                textBox1.Text = comboBox1.SelectedValue.ToString();
            }
            
        }
        public void Search()
        {
            ds.Clear();

           
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[SIGN] AS '簽名' ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(@" WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' ",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    dateTimePicker3.Value = Convert.ToDateTime(row.Cells["日期"].Value.ToString());
                    comboBox1.Text = row.Cells["部門名"].Value.ToString();
                    textBox1.Text = row.Cells["部門"].Value.ToString();                    
                    comboBox2.Text = row.Cells["飲品"].Value.ToString();
                    textBox2.Text = row.Cells["其他"].Value.ToString();
                    textBox3.Text = row.Cells["數量"].Value.ToString();
                    textBox4.Text = row.Cells["原因"].Value.ToString();
                    textBoxID.Text = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBoxID.Text = null;

                }
            }
        }
        public void SETSTATUS()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBoxID.Text = null;
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
          
        }
        public void SETSTATUS2()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;

        }
        public void SETSTAUSFIANL()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
        }
        public void UPDATE()
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

              
                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(" SET [DATES]='{0}',[DEP]='{1}',[DEPNAME]='{2}',[DRINK]='{3}',[OTHERS]='{4}',[CUP]='{5}',[REASON]='{6}',[SIGN]='{7}'", dateTimePicker3.Value.ToString("yyyyMMdd"), textBox1.Text, comboBox1.Text, comboBox2.Text, textBox2.Text, textBox3.Text, textBox4.Text,null);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", textBoxID.Text);
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

        public void ADD()
        {
            try
            {
                
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
            
                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(" ([DATES],[DEP],[DEPNAME],[DRINK],[OTHERS],[CUP],[REASON],[SIGN])");
                sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')",dateTimePicker3.Value.ToString("yyyyMMdd"),textBox1.Text,comboBox1.Text,comboBox2.Text,textBox2.Text,textBox3.Text,textBox4.Text,null);
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
        public void DEL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'",textBoxID.Text);
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

        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\飲品記錄表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();


            FASTSQL.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[SIGN] AS '簽名' ,[ID]");
            FASTSQL.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
            FASTSQL.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' ", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            STATUS = "ADD";
            SETSTATUS();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = "EDIT";
           
            SETSTATUS2();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("EDIT"))
            {
                UPDATE();
            }
            else if (STATUS.Equals("ADD"))
            {
                ADD();
            }

            STATUS = null;

            SETSTAUSFIANL();

            Search();
            MessageBox.Show("完成");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            STATUS = null;
            string message =  " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DEL();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            Search();
            MessageBox.Show("完成");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }


        #endregion


    }
}
