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
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
        SqlDataAdapter adapter9 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder9 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataSet ds9 = new DataSet();
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
        string TD003;
        string TD004;
        string TD005;
        string TD007;
        string TD011;

        string CHECKDEL;


        public class BOMTDDATA
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
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
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
            comboBox3load();

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

        public void comboBox3load()
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
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ME001";
            comboBox3.DisplayMember = "ME002";
            sqlConn.Close();

            textBox1.Text = comboBox1.SelectedValue.ToString();


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
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' ",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],111)");
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


                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[DRINKID] AS '品號' ,[SIGN] AS '簽名' ,MB004 AS '單位',[ID]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=[DRINKID] ");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' ", dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"   ORDER BY CONVERT(NVARCHAR,[DATES],112)");
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

        public void Search3()
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
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' AND [DEPNAME]='{2}' ", dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"),comboBox3.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],111)");
                sbSql.AppendFormat(@"  ");

                adapter9 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder9 = new SqlCommandBuilder(adapter9);
                sqlConn.Open();
                ds9.Clear();
                adapter9.Fill(ds9, "ds9");
                sqlConn.Close();


                if (ds9.Tables["ds9"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds9.Tables["ds9"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds9.Tables["ds9"];
                        dataGridView4.AutoResizeColumns();


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

            FASTSQL.AppendFormat(@" SELECT [DRINK] AS '飲品' ,[OTHERS] AS '其他' ,SUM([CUP]) AS '數量' ");
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
                    TD003 = row.Cells["日期"].Value.ToString();
                    TD004 = row.Cells["品號"].Value.ToString();
                    TD005 = row.Cells["單位"].Value.ToString();
                    TD007 = row.Cells["數量"].Value.ToString();
                    TD011 = row.Cells["原因"].Value.ToString();

                    CHECKBOMTDRESLUT();
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
                    TD003 = null;
                    TD004 = null;
                    TD005 = null;
                    TD007 = null;
                    TD011 = null;

                }
            }
        }


        public void ADDBOMTDRESLUT()
        {
            if(!string.IsNullOrEmpty(SID) && !string.IsNullOrEmpty(TD001) &&!string.IsNullOrEmpty(TD002))
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
                sbSql.AppendFormat(@"  WHERE  TD001='{0}' AND TD003='{1}'", TD001, TD003.ToString());
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
                return TD003.ToString() + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TD002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return TD003.ToString() + temp.ToString();
            }

            return null;
        }
        public void CHECKBOMTDRESLUT()
        {
           
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [TD001] AS '組合單' ,[TD002]  AS '組合號',[SID] AS '來源' ");
                sbSql.AppendFormat(@" FROM [TKMK].[dbo].[BOMTDRESLUT] ");
                sbSql.AppendFormat(@" WHERE [SID]='{0}' ",textBoxID2.Text.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "ds7");
                sqlConn.Close();


                if (ds7.Tables["ds7"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds7.Tables["ds7"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds7.Tables["ds7"];
                        dataGridView3.AutoResizeColumns();


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

        public void ADDBOMTDTE()
        {
            BOMTDDATA BOMTD = new BOMTDDATA();
            BOMTD = SETBOMTD();

            if(!string.IsNullOrEmpty(TD001)&& !string.IsNullOrEmpty(TD002) && !string.IsNullOrEmpty(TD003) && !string.IsNullOrEmpty(TD004) && !string.IsNullOrEmpty(TD005) && !string.IsNullOrEmpty(TD007) )
            {
                try
                {

                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[BOMTD]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010]");
                    sbSql.AppendFormat(" ,[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020]");
                    sbSql.AppendFormat(" ,[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD029],[TD030]");
                    sbSql.AppendFormat(" ,[TD031],[TD032],[TD033],[TD034],[TD035],[TD036]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", BOMTD.COMPANY, BOMTD.CREATOR, BOMTD.USR_GROUP, BOMTD.CREATE_DATE, BOMTD.MODIFIER, BOMTD.MODI_DATE, BOMTD.FLAG, BOMTD.CREATE_TIME, BOMTD.MODI_TIME, BOMTD.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',", BOMTD.TRANS_NAME, BOMTD.sync_date, BOMTD.sync_time, BOMTD.sync_mark, BOMTD.sync_count, BOMTD.DataUser, BOMTD.DataGroup);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", BOMTD.TD001, BOMTD.TD002, BOMTD.TD003, BOMTD.TD004, BOMTD.TD005, BOMTD.TD006, BOMTD.TD007, BOMTD.TD008, BOMTD.TD009, BOMTD.TD010);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", BOMTD.TD011, BOMTD.TD012, BOMTD.TD013, BOMTD.TD014, BOMTD.TD015, BOMTD.TD016, BOMTD.TD017, BOMTD.TD018, BOMTD.TD019, BOMTD.TD020);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", BOMTD.TD021, BOMTD.TD022, BOMTD.TD023, BOMTD.TD024, BOMTD.TD025, BOMTD.TD026, BOMTD.TD027, BOMTD.TD028, BOMTD.TD029, BOMTD.TD030);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}',", BOMTD.TD031, BOMTD.TD032, BOMTD.TD033, BOMTD.TD034, BOMTD.TD035, BOMTD.TD036);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", BOMTD.UDF01, BOMTD.UDF02, BOMTD.UDF03, BOMTD.UDF04, BOMTD.UDF05, BOMTD.UDF06, BOMTD.UDF07, BOMTD.UDF08, BOMTD.UDF09, BOMTD.UDF10);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[BOMTE]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010]");
                    sbSql.AppendFormat(" ,[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020]");
                    sbSql.AppendFormat(" ,[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" SELECT");
                    sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", BOMTD.COMPANY, BOMTD.CREATOR, BOMTD.USR_GROUP, BOMTD.CREATE_DATE, BOMTD.MODIFIER, BOMTD.MODI_DATE, BOMTD.FLAG, BOMTD.CREATE_TIME, BOMTD.MODI_TIME, BOMTD.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],'{4}' [sync_count],'{5}' [DataUser],'{6}' [DataGroup]", BOMTD.TRANS_NAME, BOMTD.sync_date, BOMTD.sync_time, BOMTD.sync_mark, BOMTD.sync_count, BOMTD.DataUser, BOMTD.DataGroup);
                    sbSql.AppendFormat(" ,'{0}' [TE001],'{1}' [TE002],[BOMMD].MD002 [TE003],[BOMMD].MD003 [TE004],[BOMMD].MD004 [TE005],[BOMMD].MD005 [TE006],[INVMB].MB017 [TE007],ROUND({2}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) [TE008],'{3}' [TE009],'{4}' [TE010]", TD001, TD002, TD007, null, 'N');
                    sbSql.AppendFormat(" ,'{0}' [TE011],'{1}' [TE012],'{2}' [TE013],'{3}' [TE014],'{4}' [TE015],'{5}' [TE016],'{6}' [TE017],'{7}' [TE018],'{8}' [TE019],'{9}' [TE020]", '0', '0', null, null, null, '0', '0', null, null, null);
                    sbSql.AppendFormat(" ,'{0}' [TE021],'{1}' [TE022],'{2}' [TE023],'{3}' [TE024],'{4}' [TE025],'{5}' [TE026],'{6}' [TE027],'{7}' [TE028],'{8}' [TE029]", '0', null, '0', null, '0', null, null, null, '0');
                    sbSql.AppendFormat(" ,'{0}' [UDF01],'{1}' [UDF02],'{2}' [UDF03],'{3}' [UDF04],'{4}' [UDF05],'{5}' [UDF06],'{6}' [UDF07],'{7}' [UDF08],'{8}' [UDF09],'{9}' [UDF10]", null, null, null, null, null, '0', '0', '0', '0', '0');
                    sbSql.AppendFormat(" FROM [TKMK].[dbo].[MKDRINKRECORD],[TK].dbo.[INVMB],[TK].dbo.[BOMMD]");
                    sbSql.AppendFormat(" WHERE [DRINKID]=MB001 AND [BOMMD].MD001=[INVMB].MB001");
                    sbSql.AppendFormat(" AND [ID]='{0}'", textBoxID2.Text.ToString());
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
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
           
        }

        public BOMTDDATA SETBOMTD()
        {
            BOMTDDATA BOMTD = new BOMTDDATA();
            BOMTD.COMPANY = "TK";
            BOMTD.CREATOR = "160116";
            BOMTD.USR_GROUP = "116000";
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            BOMTD.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            BOMTD.MODIFIER = "160116";
            BOMTD.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            BOMTD.FLAG = "0";
            BOMTD.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            BOMTD.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            BOMTD.TRANS_TYPE = "P001";
            BOMTD.TRANS_NAME = "BOMMI05";
            BOMTD.sync_date="";
            BOMTD.sync_time = "";
            BOMTD.sync_mark = "";
            BOMTD.sync_count = "0";
            BOMTD.DataUser = "";
            BOMTD.DataGroup = "116000";
            BOMTD.TD001 = TD001;
            BOMTD.TD002 = TD002;
            BOMTD.TD003 = TD003;
            BOMTD.TD004 = TD004;
            BOMTD.TD005 = TD005;
            BOMTD.TD006 = "";
            BOMTD.TD007 = TD007;
            BOMTD.TD008 = "0";
            BOMTD.TD009 = "0";
            BOMTD.TD010 = textBox12.Text;
            BOMTD.TD011 = TD011;
            BOMTD.TD012 = "N";
            BOMTD.TD013 = "0";
            BOMTD.TD014 = TD003;
            BOMTD.TD015 = "";
            BOMTD.TD016 = "N";
            BOMTD.TD017 = "";
            BOMTD.TD018 = "";
            BOMTD.TD019 = "";
            BOMTD.TD020 = "0";
            BOMTD.TD021 = "";
            BOMTD.TD022 = "0";
            BOMTD.TD023 = textBoxID2.Text;
            BOMTD.TD024 = "0";
            BOMTD.TD025 = "N";
            BOMTD.TD026 = "";
            BOMTD.TD027 = "";
            BOMTD.TD028 = "0";
            BOMTD.TD029 = "";
            BOMTD.TD030 = "0";
            BOMTD.TD031 = "";
            BOMTD.TD032 = "0";
            BOMTD.TD033 = "";
            BOMTD.TD034 = "";
            BOMTD.TD035 = "";
            BOMTD.TD036 = "0";
            BOMTD.UDF01 = "";
            BOMTD.UDF02 = "";
            BOMTD.UDF03 = "";
            BOMTD.UDF04 = "";
            BOMTD.UDF05 = "";
            BOMTD.UDF06 = "0";
            BOMTD.UDF07 = "0";
            BOMTD.UDF08 = "0";
            BOMTD.UDF09 = "0";
            BOMTD.UDF10 = "0";


            return BOMTD;
        }

        public void CHECKBOMTD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT * FROM [TK].dbo.BOMTD WHERE TD023='{0}'",textBoxID2.Text.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
 

                adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                sqlConn.Open();
                ds8.Clear();
                adapter8.Fill(ds8, "ds8");
                sqlConn.Close();


                if (ds8.Tables["ds8"].Rows.Count == 0)
                {
                    CHECKDEL = "Y";
                }
                else
                {
                    if (ds8.Tables["ds8"].Rows.Count >= 1)
                    {
                        CHECKDEL = "N";
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

        public void DELBOMTDRESLUT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKMK].[dbo].[BOMTDRESLUT]");
                sbSql.AppendFormat(" WHERE [SID]='{0}'", SID);
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
            ADDBOMTDRESLUT();
            ADDBOMTDTE();

            CHECKBOMTDRESLUT();

        }

        private void button9_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHECKBOMTD();

                if(CHECKDEL.Equals("Y"))
                {
                    DELBOMTDRESLUT();
                    CHECKBOMTDRESLUT();
                }
                else if (CHECKDEL.Equals("N"))
                {
                    MessageBox.Show("ERP還有組合單未刪除，請先刪除");
                }

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Search3();
        }

        #endregion


    }
}
