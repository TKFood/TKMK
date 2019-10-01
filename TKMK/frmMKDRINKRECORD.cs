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
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;
        string BUYNO;
        string OLDBUYNO;
        string CHECKYN = "N";

        string SID;
        string TD001;
        string TD002;


        public class MOCTADATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            public string TD001;
            public string TD002;
            public string TD003;
            public string TD004;
            public string TD005;
            public string TD006;
            public string TD007;
            public string TD008;
            public string TD009;
            public string TD010;
            public string TD011;
            public string TD012;
            public string TD013;
            public string TD014;
            public string TD015;
            public string TD016;
            public string TD017;
            public string TD018;
            public string TD019;
            public string TD020;
            public string TD021;
            public string TD022;
            public string TD023;
            public string TD024;
            public string TD025;
            public string TD026;
            public string TD027;
            public string TD028;
            public string TD029;
            public string TD030;
            public string TD031;
            public string TD032;
            public string TD033;
            public string TD034;
            public string TD035;
            public string TD036;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public class MOCTBDATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;

        }

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

            textBox1.Text= comboBox1.SelectedValue.ToString();


        }

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [ID],[DRINKNAME] FROM [TKMK].[dbo].[DRINKNAME] WHERE [USED]='Y' ORDER BY [ID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("DRINKNAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "DRINKNAME";
            sqlConn.Close();

            ID.Text = comboBox2.SelectedValue.ToString();


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

               
                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[DRINKID] AS '品號' ,[SIGN] AS '簽名' ,[ID]");
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

        public void Search2()
        {
            ds.Clear();


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[DRINKID] AS '品號' ,[SIGN] AS '簽名' ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(@" WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' ", dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds5.Tables["ds5"];
                        dataGridView2.AutoResizeColumns();


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
                sbSql.AppendFormat(" SET [DATES]='{0}',[DEP]='{1}',[DEPNAME]='{2}',[DRINK]='{3}',[OTHERS]='{4}',[CUP]='{5}',[REASON]='{6}',[SIGN]='{7}',[DRINKID]='{8}'", dateTimePicker3.Value.ToString("yyyyMMdd"), textBox1.Text, comboBox1.Text, comboBox2.Text, textBox2.Text, textBox3.Text, textBox4.Text,null,comboBox2.SelectedValue.ToString());
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
                sbSql.AppendFormat(" ([DATES],[DEP],[DEPNAME],[DRINK],[OTHERS],[CUP],[REASON],[SIGN],[DRINKID])");
                sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", dateTimePicker3.Value.ToString("yyyyMMdd"),textBox1.Text,comboBox1.Text,comboBox2.Text,textBox2.Text,textBox3.Text,textBox4.Text,null,comboBox2.SelectedValue.ToString());
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


        public void SETFASTREPORT2()
        {

            string SQL;
            Report report2 = new Report();
            report2.Load(@"REPORT\飲品記錄表加總.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report2.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;
            report2.SetParameterValue("P1", dateTimePicker4.Value.ToString("yyyyMMdd"));
            report2.SetParameterValue("P2", dateTimePicker5.Value.ToString("yyyyMMdd"));
            report2.Preview = previewControl2;
            report2.Show();

        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@" SELECT[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,SUM([CUP]) AS '數量' ");
            FASTSQL.AppendFormat(@" FROM [TKMK].[dbo].[MKDRINKRECORD]");
            FASTSQL.AppendFormat(@" WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}'", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@" GROUP BY [DRINK],[OTHERS]");
            FASTSQL.AppendFormat(@" ORDER BY [DRINK],[OTHERS]");
            FASTSQL.AppendFormat(@" ");
            FASTSQL.AppendFormat(@" ");

            return FASTSQL.ToString();
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ID.Text = comboBox2.SelectedValue.ToString();
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox5.Text = row.Cells["日期"].Value.ToString();
                    textBox6.Text = row.Cells["部門名"].Value.ToString();
                    textBox7.Text =row.Cells["飲品"].Value.ToString();
                    textBox8.Text = row.Cells["其他"].Value.ToString();
                    textBox9.Text = row.Cells["數量"].Value.ToString();
                    textBox10.Text = row.Cells["原因"].Value.ToString();
                    textBox11.Text = row.Cells["品號"].Value.ToString();
                    textBoxID2.Text = row.Cells["ID"].Value.ToString();

                    SID= row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBoxID2.Text = null;

                    SID = null;

                }
            }
        }


        public void ADDBOMTDRESLUT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                //TD001 = "A421";
                //TD002 = "20190930003";

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[BOMTDRESLUT]");
                sbSql.AppendFormat(" ([SID],[TD001],[TD002])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", SID, TD001, TD002);
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
        public string GETMAXTD002(string TD001)
        {
            string TD002;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TD002),'00000000000') AS TD002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMTD] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TD001='{0}' AND TD003='{1}'", TD001, textBox5.Text.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "ds6");
                sqlConn.Close();


                if (ds6.Tables["ds6"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds6.Tables["ds6"].Rows.Count >= 1)
                    {
                        TD002 = SETTD002(ds6.Tables["ds6"].Rows[0]["TD002"].ToString());
                        return TD002;

                    }
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


        public string SETTD002(string TD002)
        {
            if (TD002.Equals("00000000000"))
            {
                return textBox5.Text.ToString() + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TD002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return textBox5.Text.ToString() + temp.ToString();
            }

            return null;
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
            SETFASTREPORT2();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Search2();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            TD001 = "A421";
            TD002 = GETMAXTD002(TD001);
            //ADDBOMTDRESLUT();
        }



        #endregion

      
    }
}
