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
using System.Runtime.InteropServices;

namespace TKMK
{
    public partial class frmGROUPSALESMODIFY : Form
    {
        private ProgressBar progressBar;
        private CancellationTokenSource cancellationTokenSource;

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

        string ID;
        string STATUS;

        public frmGROUPSALESMODIFY()
        {
            InitializeComponent();
        }
        private void frmGROUPSALESMODIFY_Load(object sender, EventArgs e)
        {
            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
        }

        #region FUNCTION    

        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT [ID],[CARCOMPANY],[PRINTS],[CPMMENTS] FROM [TKMK].[dbo].[GROUPCARCOMPANY] ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("CARCOMPANY", typeof(string));
          
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "CARCOMPANY";
            comboBox1.DisplayMember = "CARCOMPANY";
            sqlConn.Close();

            comboBox1.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT [ID],[NAME] FROM [TKMK].[dbo].[CARKIND] WHERE [VALID] IN ('Y') ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAME";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();

            comboBox2.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT [ID],[NAME] FROM [TKMK].[dbo].[GROUPKIND] WHERE VALID IN ('Y') ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "NAME";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();

            comboBox3.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void comboBox4load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT  [KINDS],[PARASNAMES],[DVALUES] FROM [TKMK].[dbo].[TBZPARAS] WHERE [KINDS]='ISEXCHANGE'
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARASNAMES", typeof(string));
         
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "PARASNAMES";
            comboBox4.DisplayMember = "PARASNAMES";
            sqlConn.Close();

            comboBox4.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }

        public void SEARCHGROUPSALES(string CREATEDATES)
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
                                    [SERNO] AS '序號'
                                    ,[TA008] AS '業務員帳號'
                                    ,[CARNAME] AS '車名'
                                    ,[CARNO] AS '車號'
                                    ,[CARKIND] AS '車種'
                                    ,[GROUPKIND]  AS '團類'
                                    ,[ISEXCHANGE] AS '兌換券'
                                    ,[CARCOMPANY] AS '來車公司'
                                    ,CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'
                                    ,CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間'
                                    ,[STATUS] AS '狀態'
                                    ,[ID]
                                    ,[CREATEDATES]
                                    FROM [TKMK].[dbo].[GROUPSALES]
                                    WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'
                                    ORDER BY CONVERT(nvarchar,[CREATEDATES],112),CONVERT(int,[SERNO]) 

                                    ", CREATEDATES);


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
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                       

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        {                          

                            //判断
                            if (dgRow.Cells["狀態"].Value.ToString().Trim().Equals("完成接團"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Blue;
                            }
                            else if (dgRow.Cells["狀態"].Value.ToString().Trim().Equals("取消預約"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Pink;
                            }
                            else if (dgRow.Cells["狀態"].Value.ToString().Trim().Equals("異常結案"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Red;
                            }
                        }
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
                    ID = row.Cells["ID"].Value.ToString();

                    STATUS = row.Cells["狀態"].Value.ToString().Trim();

                    textBox1.Text = ID;

                    textBox121.Text = row.Cells["序號"].Value.ToString();
                    textBox131.Text = row.Cells["車號"].Value.ToString();
                    textBox141.Text = row.Cells["車名"].Value.ToString();
                    textBox151.Text = row.Cells["業務員帳號"].Value.ToString();


                    comboBox1.Text = row.Cells["來車公司"].Value.ToString();
                    comboBox2.Text = row.Cells["車種"].Value.ToString();   
                    comboBox3.Text = row.Cells["團類"].Value.ToString();
                    comboBox4.Text = row.Cells["兌換券"].Value.ToString();
                }
                else
                {
                    ID = null;
                    STATUS = null;
                }
            }
        }

        public void UPDATE_GROUPSALES(
            string ID,
            string TA008,
            string CARNAME,
            string CARNO,
            string CARKIND,
            string GROUPKIND,
            string ISEXCHANGE,
            string CARCOMPANY
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
                                    UPDATE [TKMK].[dbo].[GROUPSALES]
                                    SET 
                                        TA008='{1}',
                                        CARNAME='{2}',
                                        CARNO='{3}',
                                        CARKIND='{4}',
                                        GROUPKIND='{5}',
                                        ISEXCHANGE='{6}',
                                        CARCOMPANY='{7}'
                                    WHERE 
                                    ID='{0}'
                                    ",
                                    ID,
                                    TA008,
                                    CARNAME,
                                    CARNO,
                                    CARKIND,
                                    GROUPKIND,
                                    ISEXCHANGE,
                                    CARCOMPANY
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
            catch(Exception ex)
            {
                MessageBox.Show("更新失敗 "+ ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }


        public void SETNULL()
        {

            textBox1.Text ="";

            textBox121.Text = ""; 
            textBox131.Text = "";
            textBox141.Text = "";
            textBox151.Text = "";
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {   
            //查詢本日來車資料
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string ID = "";
            string TA008 = "";
            string CARNAME = "";
            string CARNO = "";
            string CARKIND = "";
            string GROUPKIND = "";
            string ISEXCHANGE = "";
            string CARCOMPANY = "";
            
            if(!string.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                ID = textBox1.Text.Trim();
                TA008 = textBox151.Text.Trim();
                CARNO  = textBox131.Text.Trim();
                CARNAME = textBox141.Text.Trim();
                CARKIND = comboBox2.Text.ToString().Trim();
                GROUPKIND = comboBox3.Text.ToString().Trim();
                ISEXCHANGE = comboBox4.Text.ToString().Trim();
                CARCOMPANY = comboBox1.Text.ToString().Trim();

                UPDATE_GROUPSALES(
                            ID,
                            TA008,
                            CARNAME,
                            CARNO,
                            CARKIND,
                            GROUPKIND,
                            ISEXCHANGE,
                            CARCOMPANY
                            );

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                //MessageBox.Show("ID "+ID+ " TA008 " + TA008 + " CARNAME " + CARNAME + " CARNO " + CARNO + " CARKIND  " + CARKIND + " GROUPKIND " + GROUPKIND + " ISEXCHANGE  " + ISEXCHANGE + " CARCOMPANY " + CARCOMPANY);
            }

        }

        #endregion


    }
}
