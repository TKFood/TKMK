﻿
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
    public partial class frmGROUPSALESBYTA008 : Form
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


        string STATUSCONTROLLER = "VIEW";
        string ID = null;
        string ACCOUNT = null;
        string ISEXCHANGE = null;
        string CARKIND = null;
        string GROUPSTARTDATES = null;
        string STARTDATES = null;
        string STARTTIMES = null;
        string STATUS = null;

        int SPECIALMNUMS = 0;
        int SPECIALMONEYS = 0;
        int SPECIALNUMSMONEYS = 0;
        int EXCHANGEMONEYS = 0;
        int EXCHANGETOTALMONEYS = 0;
        int EXCHANGESALESMMONEYS = 0;
        int COMMISSIONBASEMONEYS = 0;
        int SALESMMONEYS = 0;
        decimal COMMISSIONPCT = 0;
        int COMMISSIONPCTMONEYS = 0;
        int TOTALCOMMISSIONMONEYS = 0;
        int GUSETNUM = 0;

        int ROWSINDEX = 0;
        int COLUMNSINDEX = 0;

        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        private extern static IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        public const int WM_CLOSE = 0x10;

        public frmGROUPSALESBYTA008()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox5load();

            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;

            textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }

        #region FUNCTION
        /// <summary>
        /// 定時 每1分鐘 更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {

            if (STATUSCONTROLLER.Equals("VIEW"))
            {
                dateTimePicker1.Value = GETDBDATES();
                textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
                comboBox3load();
                label29.Text = "";
                label29.Text = "更新時間" + dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");


                MESSAGESHOW MSGSHOW = new MESSAGESHOW();
                //鎖定控制項
                this.Enabled = false;
                //顯示跳出視窗
                MSGSHOW.Show();

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                SETMONEYS();
                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                SETNUMS(dateTimePicker1.Value.ToString("yyyyMMdd"));

                //關閉跳出視窗
                MSGSHOW.Close();
                //解除鎖定
                this.Enabled = true;
            }
        }
        private void frmGROUPSALESBYTA008_FormClosed(object sender, FormClosedEventArgs e)
        {
            int NUMS = FINDSEARCHGROUPSALESNOTFINISH(dateTimePicker1.Value.ToString("yyyyMMdd"));

            if (NUMS > 0)
            {
                MessageBox.Show("仍有團務還未結案!");
            }
        }
        /// <summary>
        /// 取系統日期= 今天
        /// </summary>
        /// <returns></returns>
        public DateTime GETDBDATES()
        {
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
                                    SELECT GETDATE() AS 'DATES' 
                                    ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["DATES"].ToString());

                }
                else
                {
                    return DateTime.Now;
                }

            }
            catch
            {
                return DateTime.Now;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        /// <summary>
        /// 下拉 車種
        /// </summary>
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
            Sequel.AppendFormat(@"SELECT [ID],[NAME] FROM [TKMK].[dbo].[CARKIND] ORDER BY [ID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();

        }
        /// <summary>
        /// 下拉 團類
        /// </summary>
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
            Sequel.AppendFormat(@"SELECT [ID],[NAME] FROM [TKMK].[dbo].[GROUPKIND] ORDER BY [ID] ");
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

        }
        /// <summary>
        /// 下拉 業務員/會員
        /// </summary>
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
            Sequel.AppendFormat(@"SELECT LTRIM(RTRIM((MI001)))+' '+SUBSTRING(MI002,1,3) AS 'MI001',MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%'  AND MI001 NOT IN (SELECT [EXCHANACOOUNT] FROM [TKMK].[dbo].[GROUPSALESBYTA008] WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'  AND [STATUS]='預約接團' ) ORDER BY MI001 ", dateTimePicker1.Value.ToString("yyyyMMdd"));
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MI001", typeof(string));
            dt.Columns.Add("MI002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MI001";
            comboBox3.DisplayMember = "MI001";
            sqlConn.Close();

        }
        /// <summary>
        /// 下拉 來車公司
        /// </summary>
        public void comboBox5load()
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
            Sequel.AppendFormat(@"SELECT [PARASNAMES],[DVALUES] FROM [TKMK].[dbo].[TBZPARAS] WHERE [KINDS]='CARCOMPANY' ORDER BY [PARASNAMES]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARASNAMES", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "PARASNAMES";
            comboBox5.DisplayMember = "PARASNAMES";
            sqlConn.Close();

        }

        /// <summary>
        /// 下拉 業務員/會員，文字框更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SEARCHWSCMI(comboBox3.Text.Trim().Substring(0, 7).ToString());
        }

        /// <summary>
        /// 尋找 業務員/會員
        /// </summary>
        /// <param name="MI001"></param>
        public void SEARCHWSCMI(string MI001)
        {
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
                                    SELECT MI001,SUBSTRING(MI002,1,3) AS MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%' AND MI001='{0}' ORDER BY MI001 
                                    ", MI001);
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    textBox144.Text = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        textBox144.Text = ds.Tables["TEMPds1"].Rows[0]["MI002"].ToString();

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
        /// <summary>
        /// 自動編 流水號
        /// </summary>
        /// <param name="CREATEDATES"></param>
        /// <returns></returns>
        public string FINDSERNO(string CREATEDATES)
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(MAX(SERNO),'0') SERNO FROM  [TKMK].[dbo].[GROUPSALESBYTA008] WHERE CONVERT(NVARCHAR,[CREATEDATES],112)='{0}'"
                                    , CREATEDATES);
                sbSql.AppendFormat(@"  ");  

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        string SERNO = SETSERNO(ds1.Tables["ds1"].Rows[0]["SERNO"].ToString());
                        return SERNO;

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
        /// <summary>
        /// 格式化 流水號
        /// </summary>
        /// <param name="TEMPSERNO"></param>
        /// <returns></returns>
        public string SETSERNO(string TEMPSERNO)
        {
            if (TEMPSERNO.Equals("0"))
            {
                return "1";
            }

            else
            {
                int serno = Convert.ToInt16(TEMPSERNO);
                serno = serno + 1;
                return serno.ToString();
            }
        }
        /// <summary>
        /// 找出團務資料
        /// </summary>
        /// <param name="CREATEDATES"></param>
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
                                    ,[CARNAME] AS '車名'
                                    ,[CARNO] AS '車號'
                                    ,[CARKIND] AS '車種'
                                    ,[GROUPKIND]  AS '團類'
                                    ,[ISEXCHANGE] AS '兌換券'
                                    ,[EXCHANGETOTALMONEYS] AS '券總額'
                                    ,[EXCHANGESALESMMONEYS] AS '券消費'
                                    ,[SALESMMONEYS] AS '消費總額'
                                    ,[SPECIALMNUMS] AS '特賣數'
                                    ,[SPECIALMONEYS] AS '特賣獎金'
                                    ,[COMMISSIONBASEMONEYS] AS '茶水費'
                                    ,[COMMISSIONPCTMONEYS] AS '消費獎金'
                                    ,[TOTALCOMMISSIONMONEYS] AS '總獎金'
                                    ,[CARNUM] AS '車數'
                                    ,[GUSETNUM] AS '交易筆數'
                                    ,[CARCOMPANY] AS '來車公司'
                                    ,[TA008NO] AS '業務員名'
                                    ,[TA008] AS '業務員帳號'
                                    ,[EXCHANNO] AS '優惠券名'
                                    ,[EXCHANACOOUNT] AS '優惠券帳號'
                                    ,CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'
                                    ,CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間'
                                    ,[STATUS] AS '狀態'
                                    ,CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間'
                                    ,CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間'
                                    ,[EXCHANGEMONEYS] AS '領券額'
                                    ,[ID],[CREATEDATES]
                                    FROM [TKMK].[dbo].[GROUPSALESBYTA008]
                                    WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'
                                    AND [STATUS]<>'取消預約'
                                    ORDER BY CONVERT(nvarchar,[CREATEDATES],112),CONVERT(int,[SERNO]) DESC
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
                        dataGridView1.Columns["序號"].Width = 30;
                        dataGridView1.Columns["車名"].Width = 80;
                        dataGridView1.Columns["車號"].Width = 100;
                        dataGridView1.Columns["車種"].Width = 40;
                        dataGridView1.Columns["團類"].Width = 80;
                        dataGridView1.Columns["兌換券"].Width = 20;

                        dataGridView1.Columns["券總額"].Width = 60;
                        dataGridView1.Columns["券消費"].Width = 60;
                        dataGridView1.Columns["消費總額"].Width = 80;
                        dataGridView1.Columns["特賣數"].Width = 60;
                        dataGridView1.Columns["特賣獎金"].Width = 60;
                        dataGridView1.Columns["茶水費"].Width = 60;
                        dataGridView1.Columns["消費獎金"].Width = 60;
                        dataGridView1.Columns["總獎金"].Width = 60;
                        dataGridView1.Columns["車數"].Width = 60;
                        dataGridView1.Columns["交易筆數"].Width = 60;
                        dataGridView1.Columns["業務員名"].Width = 80;
                        dataGridView1.Columns["業務員帳號"].Width = 80;
                        dataGridView1.Columns["優惠券名"].Width = 80;
                        dataGridView1.Columns["優惠券帳號"].Width = 80;
                        dataGridView1.Columns["實際到達時間"].Width = 160;

                        dataGridView1.Columns["實際離開時間"].Width = 160;
                        dataGridView1.Columns["狀態"].Width = 160;
                        dataGridView1.Columns["預計到達時間"].Width = 100;
                        dataGridView1.Columns["預計離開時間"].Width = 80;
                        //dataGridView1.Columns["抽佣比率"].Width = 80;
                        dataGridView1.Columns["領券額"].Width = 80;
                        dataGridView1.Columns["ID"].Width = 100;
                        dataGridView1.Columns["CREATEDATES"].Width = 80;

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        {
                            dgRow.Cells["車名"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["車號"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["券總額"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["券消費"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["消費總額"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["消費獎金"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["特賣數"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["特賣獎金"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["茶水費"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["總獎金"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["交易筆數"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["優惠券名"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["業務員名"].Style.Font = new Font("Tahoma", 14);                               

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


                if (ROWSINDEX > 0 || COLUMNSINDEX > 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[ROWSINDEX].Cells[COLUMNSINDEX];

                    DataGridViewRow row = dataGridView1.Rows[ROWSINDEX];
                    ID = row.Cells["ID"].Value.ToString();

                    STATUS = row.Cells["狀態"].Value.ToString().Trim();

                    textBox121.Text = row.Cells["序號"].Value.ToString();
                    textBox131.Text = row.Cells["車號"].Value.ToString();
                    textBox141.Text = row.Cells["車名"].Value.ToString();
                    textBox142.Text = row.Cells["車數"].Value.ToString();
                    textBox143.Text = row.Cells["交易筆數"].Value.ToString();
                    //textBox144.Text = row.Cells["優惠券名"].Value.ToString();
                    textBox144.Text = row.Cells["業務員名"].Value.ToString();

                    comboBox1.Text = row.Cells["車種"].Value.ToString();
                    comboBox2.Text = row.Cells["團類"].Value.ToString();
                    //comboBox3.Text = row.Cells["優惠券帳號"].Value.ToString() + ' ' + row.Cells["優惠券名"].Value.ToString();
                    comboBox3.Text = row.Cells["業務員帳號"].Value.ToString() + ' ' + row.Cells["業務員名"].Value.ToString();
                    comboBox6.Text = row.Cells["兌換券"].Value.ToString();

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

                if (dataGridView1.CurrentCell.RowIndex > 0 || dataGridView1.CurrentCell.ColumnIndex > 0)
                {
                    textBox1.Text = dataGridView1.CurrentCell.RowIndex.ToString();
                    ROWSINDEX = dataGridView1.CurrentCell.RowIndex;
                    COLUMNSINDEX = dataGridView1.CurrentCell.ColumnIndex;

                    rowindex = ROWSINDEX;
                }



                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();

                    STATUS = row.Cells["狀態"].Value.ToString().Trim();

                    textBox121.Text = row.Cells["序號"].Value.ToString();
                    textBox131.Text = row.Cells["車號"].Value.ToString();
                    textBox141.Text = row.Cells["車名"].Value.ToString();
                    textBox142.Text = row.Cells["車數"].Value.ToString();
                    textBox143.Text = row.Cells["交易筆數"].Value.ToString();
                    //textBox144.Text = row.Cells["優惠券名"].Value.ToString();
                    textBox144.Text = row.Cells["業務員名"].Value.ToString();

                    comboBox1.Text = row.Cells["車種"].Value.ToString();
                    comboBox2.Text = row.Cells["團類"].Value.ToString();
                    //comboBox3.Text = row.Cells["優惠券帳號"].Value.ToString() + ' ' + row.Cells["優惠券名"].Value.ToString();
                    comboBox3.Text = row.Cells["業務員帳號"].Value.ToString() + ' ' + row.Cells["業務員名"].Value.ToString();
                    comboBox6.Text = row.Cells["兌換券"].Value.ToString();

                }
                else
                {
                    ID = null;
                    STATUS = null;
                }
            }
        }

        /// <summary>
        /// 尋找來車 記錄
        /// </summary>
        /// <param name="CARNO"></param>
        /// <returns></returns>
        public int SEARCHGROUPCAR(string CARNO)
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
                                    SELECT [CARNO],[CARNAME],[CARKIND]
                                    FROM [TKMK].[dbo].[GROUPCAR]
                                    WHERE [CARNO]='{0}'
                                        ", CARNO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows.Count;

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        /// <summary>
        /// 新增團務 資料
        /// </summary>
        /// <param name="ID"></param>
        /// <param name="CREATEDATES"></param>
        /// <param name="SERNO"></param>
        /// <param name="CARCOMPANY"></param>
        /// <param name="TA008NO"></param>
        /// <param name="TA008"></param>
        /// <param name="CARNO"></param>
        /// <param name="CARNAME"></param>
        /// <param name="CARKIND"></param>
        /// <param name="GROUPKIND"></param>
        /// <param name="ISEXCHANGE"></param>
        /// <param name="EXCHANGEMONEYS"></param>
        /// <param name="EXCHANGETOTALMONEYS"></param>
        /// <param name="EXCHANGESALESMMONEYS"></param>
        /// <param name="SPECIALMNUMS"></param>
        /// <param name="SPECIALMONEYS"></param>
        /// <param name="SALESMMONEYS"></param>
        /// <param name="COMMISSIONBASEMONEYS"></param>
        /// <param name="COMMISSIONPCT"></param>
        /// <param name="COMMISSIONPCTMONEYS"></param>
        /// <param name="TOTALCOMMISSIONMONEYS"></param>
        /// <param name="CARNUM"></param>
        /// <param name="GUSETNUM"></param>
        /// <param name="EXCHANNO"></param>
        /// <param name="EXCHANACOOUNT"></param>
        /// <param name="PURGROUPSTARTDATES"></param>
        /// <param name="GROUPSTARTDATES"></param>
        /// <param name="PURGROUPENDDATES"></param>
        /// <param name="GROUPENDDATES"></param>
        /// <param name="STATUS"></param>
        public void ADDGROUPSALES(
            string ID
            , string CREATEDATES
            , string SERNO
            , string CARCOMPANY
            , string TA008NO
            , string TA008
            , string CARNO
            , string CARNAME
            , string CARKIND
            , string GROUPKIND
            , string ISEXCHANGE
            , string EXCHANGEMONEYS
            , string EXCHANGETOTALMONEYS
            , string EXCHANGESALESMMONEYS
            , string SPECIALMNUMS
            , string SPECIALMONEYS
            , string SALESMMONEYS
            , string COMMISSIONBASEMONEYS
            , string COMMISSIONPCT
            , string COMMISSIONPCTMONEYS
            , string TOTALCOMMISSIONMONEYS
            , string CARNUM
            , string GUSETNUM
            , string EXCHANNO
            , string EXCHANACOOUNT
            , string PURGROUPSTARTDATES
            , string GROUPSTARTDATES
            , string PURGROUPENDDATES
            , string GROUPENDDATES
            , string STATUS
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
                                    INSERT INTO [TKMK].[dbo].[GROUPSALESBYTA008]
                                    (
                                    [CREATEDATES]
                                    ,[SERNO]
                                    ,[CARCOMPANY]
                                    ,[TA008NO]
                                    ,[TA008]
                                    ,[CARNO]
                                    ,[CARNAME]
                                    ,[CARKIND]
                                    ,[GROUPKIND]
                                    ,[ISEXCHANGE]
                                    ,[EXCHANGEMONEYS]
                                    ,[EXCHANGETOTALMONEYS]
                                    ,[EXCHANGESALESMMONEYS]
                                    ,[SPECIALMNUMS]
                                    ,[SPECIALMONEYS]
                                    ,[SALESMMONEYS]
                                    ,[COMMISSIONBASEMONEYS]
                                    ,[COMMISSIONPCT]
                                    ,[COMMISSIONPCTMONEYS]
                                    ,[TOTALCOMMISSIONMONEYS]
                                    ,[CARNUM]
                                    ,[GUSETNUM]
                                    ,[EXCHANNO]
                                    ,[EXCHANACOOUNT]
                                    ,[PURGROUPSTARTDATES]
                                    ,[GROUPSTARTDATES]
                                    ,[PURGROUPENDDATES]
                                    ,[GROUPENDDATES]
                                    ,[STATUS]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    ,'{14}'
                                    ,'{15}'
                                    ,'{16}'
                                    ,'{17}'
                                    ,'{18}'
                                    ,'{19}'
                                    ,'{20}'
                                    ,'{21}'
                                    ,'{22}'
                                    ,'{23}'
                                    ,'{24}'
                                    ,'{25}'
                                    ,'{26}'
                                    ,'{27}'
                                    ,'{28}'
                                    )
                                    ", CREATEDATES
                                    , SERNO
                                    , CARCOMPANY
                                    , TA008NO
                                    , TA008
                                    , CARNO
                                    , CARNAME
                                    , CARKIND
                                    , GROUPKIND
                                    , ISEXCHANGE
                                    , EXCHANGEMONEYS
                                    , EXCHANGETOTALMONEYS
                                    , EXCHANGESALESMMONEYS
                                    , SPECIALMNUMS
                                    , SPECIALMONEYS
                                    , SALESMMONEYS
                                    , COMMISSIONBASEMONEYS
                                    , COMMISSIONPCT
                                    , COMMISSIONPCTMONEYS
                                    , TOTALCOMMISSIONMONEYS
                                    , CARNUM
                                    , GUSETNUM
                                    , EXCHANNO
                                    , EXCHANACOOUNT
                                    , PURGROUPSTARTDATES
                                    , GROUPSTARTDATES
                                    , PURGROUPENDDATES
                                    , GROUPENDDATES
                                    , STATUS
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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        /// <summary>
        /// 更新 團務 資料
        /// </summary>
        /// <param name="ID"></param>
        /// <param name="CARCOMPANY"></param>
        /// <param name="TA008NO"></param>
        /// <param name="TA008"></param>
        /// <param name="CARNO"></param>
        /// <param name="CARNAME"></param>
        /// <param name="CARKIND"></param>
        /// <param name="GROUPKIND"></param>
        /// <param name="ISEXCHANGE"></param>
        /// <param name="CARNUM"></param>
        /// <param name="GUSETNUM"></param>
        /// <param name="EXCHANNO"></param>
        /// <param name="EXCHANACOOUNT"></param>
        /// <param name="STATUS"></param>
        public void UPDATEGROUPSALES(
                                      string ID                                    
                                    , string CARCOMPANY
                                    , string TA008NO
                                    , string TA008
                                    , string CARNO
                                    , string CARNAME
                                    , string CARKIND
                                    , string GROUPKIND
                                    , string ISEXCHANGE
                                    , string CARNUM
                                    , string GUSETNUM
                                    , string EXCHANNO
                                    , string EXCHANACOOUNT
                                    , string STATUS
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
                                    UPDATE [TKMK].[dbo].[GROUPSALESBYTA008]
                                    SET 
                                    CARCOMPANY='{1}'
                                    ,TA008NO='{2}'
                                    ,TA008='{3}'
                                    ,CARNO='{4}'
                                    ,CARNAME='{5}'
                                    ,CARKIND='{6}'
                                    ,GROUPKIND='{7}'
                                    ,ISEXCHANGE='{8}'
                                    ,CARNUM='{9}'
                                    ,GUSETNUM='{10}'
                                    ,EXCHANNO='{11}'
                                    ,EXCHANACOOUNT='{12}'
                                    ,STATUS='{13}'
                                    WHERE ID='{0}'
                                  ", ID
                                    , CARCOMPANY
                                    , TA008NO
                                    , TA008
                                    , CARNO
                                    , CARNAME
                                    , CARKIND
                                    , GROUPKIND
                                    , ISEXCHANGE
                                    , CARNUM
                                    , GUSETNUM
                                    , EXCHANNO
                                    , EXCHANACOOUNT
                                    , STATUS
                                  );

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

        /// <summary>
        /// 新增來車 記錄
        /// </summary>
        /// <param name="CARNO"></param>
        /// <param name="CARNAME"></param>
        /// <param name="CARKIND"></param>
        public void ADDGROUPCAR(string CARNO, string CARNAME, string CARKIND)
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
                                    INSERT INTO [TKMK].[dbo].[GROUPCAR]
                                    ([CARNO],[CARNAME],[CARKIND])
                                    VALUES
                                    ('{0}','{1}','{2}')
                                    ", CARNO, CARNAME, CARKIND);


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

        /// <summary>
        /// 更新 記錄
        /// </summary>
        /// <param name="CARNO"></param>
        /// <param name="CARNAME"></param>
        /// <param name="CARKIND"></param>
        public void UPDATEGROUPCAR(string CARNO, string CARNAME, string CARKIND)
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
                                    UPDATE [TKMK].[dbo].[GROUPCAR]
                                    SET [CARNAME]='{1}',[CARKIND]='{2}'
                                    WHERE [CARNO]='{0}'", CARNO, CARNAME, CARKIND);
       

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
        /// <summary>
        /// 更新 業務員/會員到POS機
        /// </summary>
        /// <param name="MI001"></param>
        /// <param name="NAME"></param>
        public void UPDATETKWSCMI(string MI001, string NAME)
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


                sbSql.AppendFormat(" UPDATE [TK].[dbo].[WSCMI] SET [MI002]='{0}' WHERE MI001='{1}'", NAME, MI001);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.138' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.135' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.134' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.133' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.132' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.130' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.137' AND MI001 ='{0}'", MI001);
                sbSql.AppendFormat(" UPDATE  [TK].[dbo].[LOG_WSCMI] SET sync_mark = 'N', sync_count=0 WHERE store_ip='192.168.1.131' AND MI001 ='{0}'", MI001);
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
        /// <summary>
        /// 更新 團務的接團
        /// </summary>
        /// <param name="ID"></param>
        /// <param name="STATUS"></param>
        public void UPDATEGROUPSALESCOMPELETE_STATUS(string ID, string STATUS)
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
                                    UPDATE [TKMK].[dbo].[GROUPSALESBYTA008]
                                    SET STATUS='{1}'
                                    WHERE [ID]='{0}'
                                    ", ID, STATUS);

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

        public void GROUPSALES_UPDATE_GROUPSTARTDATES(string ID, string GROUPSTARTDATES)
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
                                    UPDATE  [TKMK].[dbo].[GROUPSALESBYTA008]
                                    SET GROUPSTARTDATES='{1}'
                                    WHERE ID='{0}'
                                    ", ID, GROUPSTARTDATES);

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
        public void GROUPSALES_UPDATE_GROUPENDDATES(string ID, string GROUPENDDATES)
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
                                    UPDATE  [TKMK].[dbo].[GROUPSALESBYTA008]
                                    SET GROUPENDDATES='{1}'
                                    WHERE ID='{0}'
                                    ", ID, GROUPENDDATES);

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

        public void SETMONEYS()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    //判断
                    if (dr.Cells["狀態"].Value.ToString().Trim().Equals("預約接團"))
                    {
                        //清空值
                        ID = null;
                        STATUSCONTROLLER = "VIEW";
                        ACCOUNT = null;
                        ISEXCHANGE = null;
                        CARKIND = null;
                        GROUPSTARTDATES = null;
                        STARTDATES = null;
                        STARTTIMES = null;
                        SPECIALMNUMS = 0;
                        SPECIALMONEYS = 0;
                        EXCHANGEMONEYS = 0;
                        EXCHANGETOTALMONEYS = 0;
                        EXCHANGESALESMMONEYS = 0;
                        COMMISSIONBASEMONEYS = 0;
                        COMMISSIONPCT = 0;
                        COMMISSIONPCTMONEYS = 0;
                        SALESMMONEYS = 0;
                        GUSETNUM = 0;
                        TOTALCOMMISSIONMONEYS = 0;

                        //依每筆資料找出key值
                        ID = dr.Cells["ID"].Value.ToString().Trim();
                        ACCOUNT = dr.Cells["優惠券帳號"].Value.ToString().Trim();
                        ISEXCHANGE = dr.Cells["兌換券"].Value.ToString().Trim();
                        CARKIND = dr.Cells["車種"].Value.ToString().Trim();
                        GROUPSTARTDATES = dr.Cells["實際到達時間"].Value.ToString().Trim();
                        STARTDATES = GROUPSTARTDATES.Substring(0, 10).Replace("-", "").ToString();
                        STARTTIMES = GROUPSTARTDATES.Substring(11, 8);

                        //DateTime dt1 = DateTime.Now;

                        //找出各項金額    
                        //SPECIALMNUMS，算出特賣品的銷貨數量，只要VALID='Y'，就計算
                        //SPECIALNUMSMONEYS，算出特賣品 組的金額，重復SPECIALMONEYS，先不用
                        //SPECIALMONEYS，算出特賣品，銷售數量/每組*組數獎金 的金額，只要VALID='Y'，就計算
                        //SALESMMONEYS，算出該會員所有銷售金額，扣掉特賣品不合併計算的總金額，AND TB010  NOT IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT] WHERE [VALID]='Y' AND [SPLITCAL]='Y') 
                        //SPECIALNUMSMONEYS = FINDSPECIALNUMSMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                        SPECIALMNUMS = FINDSPECIALMNUMS(ACCOUNT, STARTDATES, STARTTIMES);
                        SPECIALMONEYS = FINDSPECIALMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                        SALESMMONEYS = FINDSALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);

                        //兌換券金額條件判斷
                        EXCHANGESALESMMONEYS = FINDEXCHANGESALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);

                        if (ISEXCHANGE.Trim().Equals("是"))
                        {
                            int CARNUM = Convert.ToInt32(dr.Cells["車數"].Value.ToString().Trim());
                            EXCHANGEMONEYS = FINDEXCHANGEMONEYS();
                            EXCHANGETOTALMONEYS = EXCHANGEMONEYS * CARNUM;
                            //EXCHANGESALESMMONEYS = FINDEXCHANGESALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                            COMMISSIONBASEMONEYS = 0;

                            if (EXCHANGESALESMMONEYS > 0)
                            {
                                if (SALESMMONEYS > EXCHANGETOTALMONEYS)
                                {
                                    SALESMMONEYS = SALESMMONEYS - EXCHANGETOTALMONEYS;
                                }
                            }


                        }
                        else
                        {
                            EXCHANGEMONEYS = 0;
                            EXCHANGETOTALMONEYS = 0;
                            EXCHANGESALESMMONEYS = 0;

                            //COMMISSIONBASEMONEYS，找出車子的基本輔助金額
                            COMMISSIONBASEMONEYS = FINDBASEMONEYS(CARKIND);
                        }



                        //SALESMMONEYS = SALESMMONEYS - SPECIALMONEYS;
                        COMMISSIONPCT = FINDCOMMISSIONPCT(CARKIND, SALESMMONEYS);
                        COMMISSIONPCTMONEYS = Convert.ToInt32(COMMISSIONPCT * SALESMMONEYS);
                        GUSETNUM = FINDGUSETNUM(ACCOUNT, STARTDATES, STARTTIMES);
                        TOTALCOMMISSIONMONEYS = Convert.ToInt32(SPECIALMONEYS + COMMISSIONBASEMONEYS + (COMMISSIONPCT * (SALESMMONEYS)));

                        UPDATEGROUPPRODUCT(ID, EXCHANGEMONEYS.ToString(), EXCHANGETOTALMONEYS.ToString(), EXCHANGESALESMMONEYS.ToString(), SALESMMONEYS.ToString(), SPECIALMNUMS.ToString(), SPECIALMONEYS.ToString(), COMMISSIONBASEMONEYS.ToString(), COMMISSIONPCT.ToString(), COMMISSIONPCTMONEYS.ToString(), TOTALCOMMISSIONMONEYS.ToString(), GUSETNUM.ToString());
                        //DateTime dt2 = DateTime.Now;

                        //MessageBox.Show(dt1.ToString("HH:mm:ss")+"-"+ dt2.ToString("HH:mm:ss"));
                    }

                }
            }

        }

        public int FINDSPECIALMNUMS(string TA009, string TA001, string TA005)
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
                                    SELECT SUM(NUMS) AS SPECIALMNUMS
                                    FROM (
                                    SELECT [ID],[NAME],[NUM],[MONEYS],[SPLITCAL],[VALID]
                                    ,(SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006 AND TB010=ID  AND TA009='{0}' AND TA001='{1}' AND TA005>='{2}' AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES]) ) AS 'NUMS'
                                    ,((SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006 AND TB010=ID  AND TA009='{0}' AND TA001='{1}' AND TA005>='{2}' AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES]) )/[NUM]) AS 'BASENUMS'
                                    ,((SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006 AND TB010=ID  AND TA009='{0}' AND TA001='{1}' AND TA005>='{2}' AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES]) )/[NUM])*[MONEYS] AS 'SPECIALMONEYS'
                                    FROM [TKMK].[dbo].[GROUPPRODUCT]
                                    WHERE [VALID]='Y' 
                                    ) AS TEMP
                                    ", TA009, TA001, TA005);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["SPECIALMNUMS"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public int FINDSPECIALMONEYS(string TA009, string TA001, string TA005)
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
                                    SELECT SUM(SPECIALMONEYS) AS SPECIALMONEYS
                                    FROM (
                                    SELECT [ID],[NAME],[NUM],[MONEYS],[SPLITCAL],[VALID]
                                    ,(SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006 AND TB010=ID  AND TA009='{0}' AND TA001='{1}' AND TA005>='{2}' AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES]) ) AS 'NUMS'
                                    ,((SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006 AND TB010=ID  AND TA009='{0}' AND TA001='{1}' AND TA005>='{2}' AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES]) )/[NUM]) AS 'BASENUMS'
                                    ,((SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006 AND TB010=ID  AND TA009='{0}' AND TA001='{1}' AND TA005>='{2}' AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES]) )/[NUM])*[MONEYS] AS 'SPECIALMONEYS'
                                    FROM [TKMK].[dbo].[GROUPPRODUCT]
                                    WHERE [VALID]='Y' 
                                    ) AS TEMP

                                     ", TA009, TA001, TA005);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["SPECIALMONEYS"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }


        public int FINDSALESMMONEYS(string TA009, string TA001, string TA005)
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

                //將特買組的銷售金額扣掉 TB010  NOT IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT] WHERE [SPLITCAL]='Y') 
                sbSql.AppendFormat(@"  
                                    SELECT CONVERT(INT,ISNULL(SUM(TB033),0),0) AS 'SALESMMONEYS'
                                    FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)
                                    WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006  
                                    AND TB010  NOT IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT] WHERE [VALID]='Y' AND [SPLITCAL]='Y')              
                                    AND TA009='{0}'
                                    AND TA001='{1}'
                                    AND TA005>='{2}'
                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES])
                                    ", TA009, TA001, TA005);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["SALESMMONEYS"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public int FINDEXCHANGESALESMMONEYS(string TA009, string TA001, string TA005)
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
                                    SELECT CONVERT(INT,ISNULL(SUM(TA017),0)) AS EXCHANGESALESMMONEYS
                                    FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTC WITH (NOLOCK)
                                    WHERE TA001=TC001 AND TA002=TC002 AND TA003=TC003  AND TA006=TC006
                                    AND TC008='0009'
                                    AND TA009='{0}'
                                    AND TA001='{1}'
                                    AND TA005>='{2}'
                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES])
                                    ", TA009, TA001, TA005);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["EXCHANGESALESMMONEYS"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public int FINDEXCHANGEMONEYS()
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
                                    CONVERT(INT,[EXCHANGEMONEYS],0) AS EXCHANGEMONEYS   
                                    FROM [TKMK].[dbo].[GROUPEXCHANGEMONEYS]
                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["EXCHANGEMONEYS"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public int FINDBASEMONEYS(string NAME)
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
                                    SELECT CONVERT(INT,[BASEMONEYS],0) AS 'BASEMONEYS' 
                                    FROM [TKMK].[dbo].[GROUPBASE] 
                                    WHERE [NAME]='{0}'"
                                    , NAME);
        

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["BASEMONEYS"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public decimal FINDCOMMISSIONPCT(string CARKIND, int MONEYS)
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
                                    SELECT [ID],[PCTMONEYS],[NAME],[PCT]
                                    FROM [TKMK].[dbo].[GROUPPCT]
                                    WHERE [NAME]='{0}' AND ({1}-[PCTMONEYS])>=0
                                    ORDER BY ({1}-[PCTMONEYS])
                                    ", CARKIND, MONEYS);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToDecimal(ds1.Tables["ds1"].Rows[0]["PCT"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public int FINDGUSETNUM(string TA009, string TA001, string TA005)
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
                                    SELECT COUNT(TA009) AS 'GUSETNUM'
                                    FROM [TK].dbo.POSTA WITH (NOLOCK)
                                    WHERE TA009='{0}'
                                    AND TA001='{1}'
                                    AND TA005>='{2}'
                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES])
                                    ", TA009, TA001, TA005);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["GUSETNUM"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEGROUPPRODUCT(string ID, string EXCHANGEMONEYS, string EXCHANGETOTALMONEYS, string EXCHANGESALESMMONEYS, string SALESMMONEYS, string SPECIALMNUMS, string SPECIALMONEYS, string COMMISSIONBASEMONEYS, string COMMISSIONPCT, string COMMISSIONPCTMONEYS, string TOTALCOMMISSIONMONEYS, string GUSETNUM)
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
                                    UPDATE [TKMK].[dbo].[GROUPSALESBYTA008]
                                    SET [EXCHANGEMONEYS]={1},[EXCHANGETOTALMONEYS]={2},[EXCHANGESALESMMONEYS]={3},[SALESMMONEYS]={4},[SPECIALMNUMS]={5},[SPECIALMONEYS]={6},[COMMISSIONBASEMONEYS]={7},[COMMISSIONPCT]={8},[COMMISSIONPCTMONEYS]={9},[TOTALCOMMISSIONMONEYS]={10},[GUSETNUM]={11}
                                    WHERE [ID]='{0}'
                                    ", ID, EXCHANGEMONEYS, EXCHANGETOTALMONEYS, EXCHANGESALESMMONEYS, SALESMMONEYS, SPECIALMNUMS, SPECIALMONEYS, COMMISSIONBASEMONEYS, COMMISSIONPCT, COMMISSIONPCTMONEYS, TOTALCOMMISSIONMONEYS, GUSETNUM);

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

        public void SETNUMS(string GROUPSTARTDATES)
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
                                    SELECT COUNT(CARNO) AS NUMS  
                                    ,(SELECT SUM(GUSETNUM) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) ) AS GUSETNUMS
                                    ,(SELECT SUM(SALESMMONEYS) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) ) AS SALESMMONEYS
                                    ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) AND STATUS='預約接團') AS CARNUM1
                                    ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) AND STATUS='取消預約') AS CARNUM2
                                    ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) AND STATUS='異常結案') AS CARNUM3
                                    ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) AND STATUS='完成接團') AS CARNUM4
                                    ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALESBYTA008] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALESBYTA008].GROUPSTARTDATES,112) AND STATUS='預約接團') AS CARNUM5
                                    FROM [TKMK].[dbo].[GROUPSALESBYTA008] WITH(NOLOCK)
                                    WHERE CONVERT(NVARCHAR,GROUPSTARTDATES,112)='{0}'
                                    AND STATUS IN ('預約接團','完成接團')
                                    GROUP BY CONVERT(NVARCHAR,GROUPSTARTDATES,112)
                                    ", GROUPSTARTDATES);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    label12.Text = ds1.Tables["ds1"].Rows[0]["NUMS"].ToString().Trim();
                    label14.Text = ds1.Tables["ds1"].Rows[0]["GUSETNUMS"].ToString().Trim();
                    label16.Text = ds1.Tables["ds1"].Rows[0]["SALESMMONEYS"].ToString().Trim();
                    label18.Text = ds1.Tables["ds1"].Rows[0]["CARNUM1"].ToString().Trim();
                    label23.Text = ds1.Tables["ds1"].Rows[0]["CARNUM2"].ToString().Trim();
                    label20.Text = ds1.Tables["ds1"].Rows[0]["CARNUM3"].ToString().Trim();
                    label24.Text = ds1.Tables["ds1"].Rows[0]["CARNUM4"].ToString().Trim();
                    label21.Text = ds1.Tables["ds1"].Rows[0]["CARNUM5"].ToString().Trim();

                }
                else
                {
                    label12.Text = "0";
                    label14.Text = "0";
                    label16.Text = "0";
                    label18.Text = "0";
                    label23.Text = "0";
                    label20.Text = "0";
                    label21.Text = "0";
                    label24.Text = "0";

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

        public int FINDSEARCHGROUPSALESNOTFINISH(string CREATEDATES)
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
                                    SELECT COUNT([CARNO]) AS NUMS 
                                    FROM [TKMK].[dbo].[GROUPSALESBYTA008]
                                    WHERE [STATUS]='預約接團' AND CONVERT(nvarchar,[CREATEDATES],112)='{0}' 
                                    ", CREATEDATES);
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["NUMS"].ToString());
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void SETTEXT1()
        {
            textBox131.Text = null;
            textBox141.Text = null;
            textBox142.Text = "1";
            textBox143.Text = "1";

            textBox131.ReadOnly = false;
            textBox141.ReadOnly = false;
            textBox142.ReadOnly = false;
            textBox143.ReadOnly = false;

            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox5.Enabled = true;
            comboBox6.Enabled = true;
        }

        public void SETTEXT2()
        {
            textBox131.ReadOnly = true;
            textBox141.ReadOnly = true;
            textBox142.ReadOnly = true;
            textBox143.ReadOnly = true;

            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            comboBox5.Enabled = false;
            comboBox6.Enabled = false;
        }

        public void SETTEXT3()
        {
            textBox131.ReadOnly = false;
            textBox141.ReadOnly = false;
            textBox142.ReadOnly = false;
            textBox143.ReadOnly = false;

            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox5.Enabled = true;
            comboBox6.Enabled = true;

        }

        public void SETTEXT4()
        {
            textBox131.ReadOnly = true;
            textBox141.ReadOnly = true;
            textBox142.ReadOnly = true;
            textBox143.ReadOnly = true;

            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox5.Enabled = false;
            comboBox6.Enabled = false;
        }
        public void SETTEXT5()
        {
            comboBox3.Enabled = true;
        }

        public void SETTEXT6()
        {
            comboBox3.Enabled = false;
        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            MESSAGESHOW MSGSHOW = new MESSAGESHOW();
            //鎖定控制項
            this.Enabled = false;
            //顯示跳出視窗
            MSGSHOW.Show();
            //查詢本日來車資料
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            //計算佣金
            SETMONEYS();
            //查詢本日來車資料
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            //查詢本日的合計
            SETNUMS(dateTimePicker1.Value.ToString("yyyyMMdd"));

            label29.Text = "";
            label29.Text = "更新時間" + dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");



            //關閉跳出視窗
            MSGSHOW.Close();
            //解除鎖定
            this.Enabled = true;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            STATUSCONTROLLER = "ADD";

            SETTEXT1();
            comboBox3load();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            if (STATUSCONTROLLER.Equals("ADD"))
            {
                string ID = Guid.NewGuid().ToString();
                string CREATEDATES = dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string SERNO = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
                string CARNO = textBox131.Text.Trim();
                string CARNAME = textBox141.Text.Trim();
                string CARKIND = comboBox1.Text.Trim();
                string GROUPKIND = comboBox2.Text.Trim();
                string ISEXCHANGE = comboBox6.Text.Trim();



                string EXCHANGEMONEYS = "0";
                string EXCHANGETOTALMONEYS = "0";
                string EXCHANGESALESMMONEYS = "0";
                string SALESMMONEYS = "0";
                string SPECIALMNUMS = "0";
                string SPECIALMONEYS = "0";
                string COMMISSIONBASEMONEYS = "0";
                string COMMISSIONPCT = "0";
                string COMMISSIONPCTMONEYS = "0";
                string TOTALCOMMISSIONMONEYS = "0";
                string CARNUM = textBox142.Text.Trim();
                string GUSETNUM = textBox143.Text.Trim();
                string EXCHANNO = textBox144.Text.Trim();
                string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString();
                string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string GROUPENDDATES = "1911/1/1";
                string STATUS = "預約接團";
                string TA008 = comboBox3.Text.Trim().Substring(0, 7).ToString();
                string TA008NO = textBox144.Text.Trim();
                string CARCOMPANY = comboBox5.SelectedValue.ToString();

                try
                {
                    if (!string.IsNullOrEmpty(SERNO) && !string.IsNullOrEmpty(CARNO) && !string.IsNullOrEmpty(EXCHANNO) && !string.IsNullOrEmpty(EXCHANACOOUNT) && Convert.ToInt32(CARNUM) >= 1)
                    {
                        ADDGROUPSALES(
                        ID
                        , CREATEDATES
                        , SERNO
                        , CARCOMPANY
                        , TA008NO
                        , TA008
                        , CARNO
                        , CARNAME
                        , CARKIND
                        , GROUPKIND
                        , ISEXCHANGE
                        , EXCHANGEMONEYS
                        , EXCHANGETOTALMONEYS
                        , EXCHANGESALESMMONEYS
                        , SPECIALMNUMS
                        , SPECIALMONEYS
                        , SALESMMONEYS
                        , COMMISSIONBASEMONEYS
                        , COMMISSIONPCT
                        , COMMISSIONPCTMONEYS
                        , TOTALCOMMISSIONMONEYS
                        , CARNUM
                        , GUSETNUM
                        , EXCHANNO
                        , EXCHANACOOUNT
                        , PURGROUPSTARTDATES
                        , GROUPSTARTDATES
                        , PURGROUPENDDATES
                        , GROUPENDDATES
                        , STATUS
                       );

                        textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
                        SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                    }
                    else
                    {
                        MessageBox.Show("團務資料少填 或車數 的填寫有問題");
                    }
                }
                catch
                {
                    MessageBox.Show("團務資料少填 或 車數 的填寫有問題");
                }
                finally
                {

                }


                if (!string.IsNullOrEmpty(CARNO) && !string.IsNullOrEmpty(CARNAME) && !string.IsNullOrEmpty(CARKIND))
                {
                    int ISCAR = SEARCHGROUPCAR(CARNO);

                    if (ISCAR == 0)
                    {
                        ADDGROUPCAR(CARNO, CARNAME, CARKIND);
                    }
                    else if (ISCAR == 1)
                    {
                        UPDATEGROUPCAR(CARNO, CARNAME, CARKIND);
                    }
                }

                //UPDATETKWSCMI(EXCHANACOOUNT, EXCHANNO + ' ' + CARNAME);
            }
            else if (STATUSCONTROLLER.Equals("EDIT"))
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    string CARNO = textBox131.Text.Trim();
                    string CARNAME = textBox141.Text.Trim();
                    string CARKIND = comboBox1.Text.Trim();
                    string GROUPKIND = comboBox2.Text.Trim();
                    string ISEXCHANGE = comboBox6.Text.Trim();

                    string CARNUM = textBox142.Text.Trim();
                    string GUSETNUM = textBox143.Text.Trim();
                    string EXCHANNO = textBox144.Text.Trim();
                    string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString();
                    string CARCOMPANY = comboBox5.SelectedValue.ToString();
                    string TA008NO = textBox144.Text.Trim();
                    string TA008 = comboBox3.Text.Trim().Substring(0, 7).ToString();
                    //string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");

                    if (!string.IsNullOrEmpty(ID) && !string.IsNullOrEmpty(CARNO) && !string.IsNullOrEmpty(EXCHANNO) && !string.IsNullOrEmpty(EXCHANACOOUNT))
                    {                        
                         UPDATEGROUPSALES(
                                      ID
                                    , CARCOMPANY
                                    , TA008NO
                                    , TA008
                                    , CARNO
                                    , CARNAME
                                    , CARKIND
                                    , GROUPKIND
                                    , ISEXCHANGE
                                    , CARNUM
                                    , GUSETNUM
                                    , EXCHANNO
                                    , EXCHANACOOUNT
                                    , "預約接團"
                                    );
                    }

                    if (!string.IsNullOrEmpty(CARNO) && !string.IsNullOrEmpty(CARNAME) && !string.IsNullOrEmpty(CARKIND))
                    {
                        int ISCAR = SEARCHGROUPCAR(CARNO);

                        if (ISCAR == 0)
                        {
                            ADDGROUPCAR(CARNO, CARNAME, CARKIND);
                        }
                        else if (ISCAR == 1)
                        {
                            UPDATEGROUPCAR(CARNO, CARNAME, CARKIND);
                        }
                    }

                    UPDATETKWSCMI(EXCHANACOOUNT, EXCHANNO + ' ' + CARNAME);
                }



            }



            SETTEXT2();
            SETTEXT4();
            SETTEXT6();
            STATUSCONTROLLER = "VIEW";

            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SETTEXT2();
            SETTEXT4();
            SETTEXT6();
            STATUSCONTROLLER = "VIEW";

            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("預約接團"))
            {
                SETTEXT3();
                STATUSCONTROLLER = "EDIT";
            }
            else
            {
                MessageBox.Show("不是預約接團，不能修改");
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {

            if (STATUS.Equals("預約接團"))
            {
                comboBox3load();
                SETTEXT5();
                STATUSCONTROLLER = "EDIT";
            }
            else
            {
                MessageBox.Show("不是預約接團，不能修改");
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string DTIMES = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            if (!string.IsNullOrEmpty(ID))
            {
                GROUPSALES_UPDATE_GROUPSTARTDATES(ID, DTIMES);
                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string DTIMES = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            if (!string.IsNullOrEmpty(ID))
            {
                GROUPSALES_UPDATE_GROUPENDDATES(ID, DTIMES);
                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("預約接團"))
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    string CARNO = textBox131.Text.Trim();
                    string CARNAME = textBox141.Text.Trim();
                    string CARKIND = comboBox1.Text.Trim();
                    string GROUPKIND = comboBox2.Text.Trim();
                    string ISEXCHANGE = comboBox6.Text.Trim();

                    string CARNUM = textBox142.Text.Trim();
                    string GUSETNUM = textBox143.Text.Trim();
                    string EXCHANNO = textBox144.Text.Trim();
                    string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString();
                    string CARCOMPANY = comboBox5.SelectedValue.ToString();
                    string TA008NO = textBox144.Text.Trim();
                    string TA008 = comboBox3.Text.Trim().Substring(0, 7).ToString();               
                    
                    UPDATEGROUPSALES(
                                      ID
                                    , CARCOMPANY
                                    , TA008NO
                                    , TA008
                                    , CARNO
                                    , CARNAME
                                    , CARKIND
                                    , GROUPKIND
                                    , ISEXCHANGE
                                    , CARNUM
                                    , GUSETNUM
                                    , EXCHANNO
                                    , EXCHANACOOUNT
                                    , "取消預約"
                                    );
                }

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                MessageBox.Show("不是預約接團，不能修改");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("預約接團"))
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    string CARNO = textBox131.Text.Trim();
                    string CARNAME = textBox141.Text.Trim();
                    string CARKIND = comboBox1.Text.Trim();
                    string GROUPKIND = comboBox2.Text.Trim();
                    string ISEXCHANGE = comboBox6.Text.Trim();

                    string CARNUM = textBox142.Text.Trim();
                    string GUSETNUM = textBox143.Text.Trim();
                    string EXCHANNO = textBox144.Text.Trim();
                    string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString();
                    string CARCOMPANY = comboBox5.SelectedValue.ToString();
                    string TA008NO = textBox144.Text.Trim();
                    string TA008 = comboBox3.Text.Trim().Substring(0, 7).ToString();
                  
                    UPDATEGROUPSALES(
                                      ID
                                    , CARCOMPANY
                                    , TA008NO
                                    , TA008
                                    , CARNO
                                    , CARNAME
                                    , CARKIND
                                    , GROUPKIND
                                    , ISEXCHANGE
                                    , CARNUM
                                    , GUSETNUM
                                    , EXCHANNO
                                    , EXCHANACOOUNT
                                    , "異常結案"
                                    );
                }

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                MessageBox.Show("不是預約接團，不能修改");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("預約接團"))
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    //string GROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //UPDATEGROUPSALESCOMPELETE(ID, GROUPENDDATES, "完成接團");

                    UPDATEGROUPSALESCOMPELETE_STATUS(ID, "完成接團");
                }

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                MessageBox.Show("不是預約接團，不能修改");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }



        #endregion

      
    }
}
