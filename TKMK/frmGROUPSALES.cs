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
    public partial class frmGROUPSALES : Form
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

        public frmGROUPSALES()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();

            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;

            textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

            timer1.Enabled = true;
            timer1.Interval = 1000 * 30;
            timer1.Start();
        }

        #region FUNCTION
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (STATUSCONTROLLER.Equals("VIEW") )
            {
                dateTimePicker1.Value = GETDBDATES();
                textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
                comboBox3load();
                label29.Text = "";
                label29.Text = "更新時間" + dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");


                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                SETMONEYS();
                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                SETNUMS(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            

            //dateTimePicker1.Value = DateTime.Now;
            //dateTimePicker2.Value = DateTime.Now;
            //dateTimePicker3.Value = DateTime.Now;
        }
        private void frmGROUPSALES_FormClosed(object sender, FormClosedEventArgs e)
        {
            int NUMS= FINDSEARCHGROUPSALESNOTFINISH(dateTimePicker1.Value.ToString("yyyyMMdd"));

            if(NUMS>0)
            {
                MessageBox.Show("仍有團務還未結案!");
            }
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
            dateTimePicker3.Value = dateTimePicker1.Value;

            textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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

        public void comboBox3load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT LTRIM(RTRIM((MI001)))+' '+SUBSTRING(MI002,1,3) AS 'MI001',MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%'  AND MI001 NOT IN (SELECT [EXCHANACOOUNT] FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'  AND [STATUS]='預約接團' ) ORDER BY MI001 ", dateTimePicker1.Value.ToString("yyyyMMdd"));
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

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SEARCHWSCMI(comboBox3.Text.Trim().Substring(0,7).ToString());
        }


        public DateTime GETDBDATES()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT GETDATE() AS 'DATES' ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public void SEARCHWSCMI(string MI001)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MI001,SUBSTRING(MI002,1,3) AS MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%' AND MI001='{0}' ORDER BY MI001 ", MI001);
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

        public void SEARCHGROUPSALES(string CREATEDATES)
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

                sbSql.AppendFormat(@" SELECT  
                                    [SERNO] AS '序號',[CARNAME] AS '車名',[CARNO] AS '車號',[CARKIND] AS '車種',[GROUPKIND]  AS '團類',[ISEXCHANGE] AS '兌換券',[EXCHANGETOTALMONEYS] AS '券總額',[EXCHANGESALESMMONEYS] AS '券消費',[SALESMMONEYS] AS '消費總額'
                                    ,[SPECIALMNUMS] AS '特賣數',[SPECIALMONEYS] AS '特賣獎金',[COMMISSIONBASEMONEYS] AS '茶水費',[COMMISSIONPCTMONEYS] AS '消費獎金',[TOTALCOMMISSIONMONEYS] AS '總獎金',[CARNUM] AS '車數',[GUSETNUM] AS '來客數',[EXCHANNO] AS '優惠券名',[EXCHANACOOUNT] AS '優惠券帳號',CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間',CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間',[STATUS] AS '狀態'
                                    ,CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間',CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間',[COMMISSIONPCT] AS '抽佣比率',[EXCHANGEMONEYS] AS '領券額',[ID],[CREATEDATES]
                                    FROM [TKMK].[dbo].[GROUPSALES]
                                    WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'
                                    AND [STATUS]<>'取消預約'
                                    ORDER BY CONVERT(nvarchar,[CREATEDATES],112),CONVERT(int,[SERNO]) DESC ", CREATEDATES);


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
                        dataGridView1.Columns["消費總額"].Width =80;
                        dataGridView1.Columns["特賣數"].Width = 60;
                        dataGridView1.Columns["特賣獎金"].Width = 60;
                        dataGridView1.Columns["茶水費"].Width = 60;
                        dataGridView1.Columns["消費獎金"].Width = 60;
                        dataGridView1.Columns["總獎金"].Width = 60;
                        dataGridView1.Columns["車數"].Width = 60;
                        dataGridView1.Columns["來客數"].Width = 60;
                        dataGridView1.Columns["優惠券名"].Width = 60;
                        dataGridView1.Columns["優惠券帳號"].Width = 80;
                        dataGridView1.Columns["實際到達時間"].Width = 160;

                        dataGridView1.Columns["實際離開時間"].Width = 160;
                        dataGridView1.Columns["狀態"].Width = 160;
                        dataGridView1.Columns["預計到達時間"].Width = 100;
                        dataGridView1.Columns["預計離開時間"].Width = 80;
                        dataGridView1.Columns["抽佣比率"].Width = 80;
                        dataGridView1.Columns["領券額"].Width = 80;
                        dataGridView1.Columns["ID"].Width = 200;
                        dataGridView1.Columns["CREATEDATES"].Width = 80;

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        {
                            dgRow.Cells["車名"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["車號"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["券總額"].Style.Font= new Font("Tahoma", 14);
                            dgRow.Cells["券消費"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["消費總額"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["消費獎金"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["特賣數"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["特賣獎金"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["茶水費"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["總獎金"].Style.Font = new Font("Tahoma", 14);
                            dgRow.Cells["來客數"].Style.Font = new Font("Tahoma", 14);

                            //判断
                            if (dgRow.Cells[20].Value.ToString().Trim().Equals("完成接團"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Blue;
                            }
                            else if (dgRow.Cells[20].Value.ToString().Trim().Equals("取消預約"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Pink;
                            }
                            else if (dgRow.Cells[20].Value.ToString().Trim().Equals("異常結案"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Red;
                            }
                        }
                    }

                }


                if(ROWSINDEX > 0|| COLUMNSINDEX>0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[ROWSINDEX].Cells[COLUMNSINDEX];

                    DataGridViewRow row = dataGridView1.Rows[ROWSINDEX];
                    ID = row.Cells["ID"].Value.ToString();

                    STATUS = row.Cells["狀態"].Value.ToString().Trim();

                    textBox121.Text = row.Cells["序號"].Value.ToString();
                    textBox131.Text = row.Cells["車號"].Value.ToString();
                    textBox141.Text = row.Cells["車名"].Value.ToString();
                    textBox142.Text = row.Cells["車數"].Value.ToString();
                    textBox143.Text = row.Cells["來客數"].Value.ToString();
                    textBox144.Text = row.Cells["優惠券名"].Value.ToString();

                    comboBox1.Text = row.Cells["車種"].Value.ToString();
                    comboBox2.Text = row.Cells["團類"].Value.ToString();
                    comboBox3.Text = row.Cells["優惠券帳號"].Value.ToString() + ' ' + row.Cells["優惠券名"].Value.ToString();

                    if (row.Cells["兌換券"].Value.ToString().Equals("Y"))
                    {
                        checkBox1.Checked = true;
                    }
                    else if (row.Cells["兌換券"].Value.ToString().Equals("N"))
                    {
                        checkBox1.Checked = false; 
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

        public void SEARCHGROUPSALES2(string CREATEDATES, string STATUS)
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

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [SERNO] AS '序號',[CARNAME] AS '車名',[CARNO] AS '車號',[CARKIND] AS '車種',[GROUPKIND]  AS '團類',[ISEXCHANGE] AS '兌換券',[EXCHANGETOTALMONEYS] AS '券總額',[EXCHANGESALESMMONEYS] AS '券消費',[SALESMMONEYS] AS '消費總額'");
                sbSql.AppendFormat(@"  ,[SPECIALMNUMS] AS '特賣數',[SPECIALMONEYS] AS '特賣獎金',[COMMISSIONBASEMONEYS] AS '茶水費',[COMMISSIONPCTMONEYS] AS '消費獎金',[TOTALCOMMISSIONMONEYS] AS '總獎金',[CARNUM] AS '車數',[GUSETNUM] AS '來客數',[EXCHANNO] AS '優惠券名',[EXCHANACOOUNT] AS '優惠券帳號',CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間',CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間',[STATUS] AS '狀態'");
                sbSql.AppendFormat(@"  ,CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間',CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間',[COMMISSIONPCT] AS '抽佣比率',[EXCHANGEMONEYS] AS '領券額',[ID],[CREATEDATES]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(@"  WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}' ", CREATEDATES);
                sbSql.AppendFormat(@"  AND [STATUS]='{0}'", STATUS);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(nvarchar,[CREATEDATES],112),CONVERT(int,[SERNO]) DESC");
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
                        dataGridView1.Columns[0].Width = 30;
                        dataGridView1.Columns[1].Width = 80;
                        dataGridView1.Columns[2].Width = 80;
                        dataGridView1.Columns[3].Width = 40;
                        dataGridView1.Columns[4].Width = 80;
                        dataGridView1.Columns[5].Width = 20;

                        dataGridView1.Columns[6].Width = 60;
                        dataGridView1.Columns[7].Width = 60;
                        dataGridView1.Columns[8].Width = 60;
                        dataGridView1.Columns[9].Width = 60;
                        dataGridView1.Columns[10].Width = 60;
                        dataGridView1.Columns[11].Width = 60;
                        dataGridView1.Columns[12].Width = 60;
                        dataGridView1.Columns[13].Width = 60;
                        dataGridView1.Columns[14].Width = 60;
                        dataGridView1.Columns[15].Width = 60;
                        dataGridView1.Columns[16].Width = 60;
                        dataGridView1.Columns[17].Width = 80;
                        dataGridView1.Columns[18].Width = 160;

                        dataGridView1.Columns[19].Width = 160;
                        dataGridView1.Columns[20].Width = 160;
                        dataGridView1.Columns[21].Width = 100;
                        dataGridView1.Columns[22].Width = 80;
                        dataGridView1.Columns[23].Width = 80;
                        dataGridView1.Columns[24].Width = 80;
                        dataGridView1.Columns[25].Width = 200;
                        dataGridView1.Columns[26].Width = 80;

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        {
                            //判断
                            if (dgRow.Cells[20].Value.ToString().Trim().Equals("完成接團"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Blue;
                            }
                            else if (dgRow.Cells[20].Value.ToString().Trim().Equals("取消預約"))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.ForeColor = Color.Pink;
                            }
                            else if (dgRow.Cells[20].Value.ToString().Trim().Equals("異常結案"))
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
                    textBox143.Text = row.Cells["來客數"].Value.ToString();
                    textBox144.Text = row.Cells["優惠券名"].Value.ToString();

                    comboBox1.Text = row.Cells["車種"].Value.ToString();
                    comboBox2.Text = row.Cells["團類"].Value.ToString();
                    comboBox3.Text = row.Cells["優惠券帳號"].Value.ToString() + ' ' + row.Cells["優惠券名"].Value.ToString();

                    if (row.Cells["兌換券"].Value.ToString().Equals("Y"))
                    {
                        checkBox1.Checked = true;
                    }
                    else if (row.Cells["兌換券"].Value.ToString().Equals("N"))
                    {
                        checkBox1.Checked = false;
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

                if(dataGridView1.CurrentCell.RowIndex>0|| dataGridView1.CurrentCell.ColumnIndex>0)
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

                    STATUS= row.Cells["狀態"].Value.ToString().Trim();

                    textBox121.Text = row.Cells["序號"].Value.ToString();
                    textBox131.Text = row.Cells["車號"].Value.ToString();
                    textBox141.Text = row.Cells["車名"].Value.ToString();
                    textBox142.Text = row.Cells["車數"].Value.ToString();
                    textBox143.Text = row.Cells["來客數"].Value.ToString();
                    textBox144.Text = row.Cells["優惠券名"].Value.ToString();

                    comboBox1.Text = row.Cells["車種"].Value.ToString();
                    comboBox2.Text = row.Cells["團類"].Value.ToString();
                    comboBox3.Text = row.Cells["優惠券帳號"].Value.ToString()+' '+ row.Cells["優惠券名"].Value.ToString();

                    if(row.Cells["兌換券"].Value.ToString().Equals("Y"))
                    {
                        checkBox1.Checked = true;
                    }
                    else if(row.Cells["兌換券"].Value.ToString().Equals("N"))
                    {
                        checkBox1.Checked = false;
                    }
                   
                }
                else
                {
                    ID = null;
                    STATUS = null;
                }
            }
        }

        public string FINDSERNO(string CREATEDATES)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

               
                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(SERNO),'0') SERNO FROM  [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(NVARCHAR,[CREATEDATES],112)='{0}'",CREATEDATES);
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
        public void ADDGROUPSALES(
            string ID, string CREATEDATES, string SERNO, string CARNO, string CARNAME, string CARKIND,string GROUPKIND, string ISEXCHANGE, string EXCHANGEMONEYS,string EXCHANGETOTALMONEYS
            , string EXCHANGESALESMMONEYS , string SALESMMONEYS , string SPECIALMNUMS, string SPECIALMONEYS, string COMMISSIONBASEMONEYS, string COMMISSIONPCTMONEYS, string TOTALCOMMISSIONMONEYS, string CARNUM, string GUSETNUM, string EXCHANNO
            , string EXCHANACOOUNT, string PURGROUPSTARTDATES , string GROUPSTARTDATES, string PURGROUPENDDATES, string GROUPENDDATES, string STATUS
            )
        {


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(" (");
                sbSql.AppendFormat(" [ID],[CREATEDATES],[SERNO],[CARNO],[CARNAME],[CARKIND],[GROUPKIND],[ISEXCHANGE],[EXCHANGEMONEYS],[EXCHANGETOTALMONEYS]");
                sbSql.AppendFormat(" ,[EXCHANGESALESMMONEYS],[SALESMMONEYS],[SPECIALMNUMS],[SPECIALMONEYS],[COMMISSIONBASEMONEYS],[COMMISSIONPCTMONEYS],[TOTALCOMMISSIONMONEYS],[CARNUM],[GUSETNUM],[EXCHANNO]");
                sbSql.AppendFormat(" ,[EXCHANACOOUNT],[PURGROUPSTARTDATES],[GROUPSTARTDATES],[PURGROUPENDDATES],[GROUPENDDATES],[STATUS]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" (");
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}',{8},{9},", ID, CREATEDATES, SERNO, CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, EXCHANGEMONEYS, EXCHANGETOTALMONEYS);
                sbSql.AppendFormat("  {0},{1},{2},{3},{4},{5},{6},{7},{8},'{9}',", EXCHANGESALESMMONEYS, SALESMMONEYS, SPECIALMNUMS, SPECIALMONEYS, COMMISSIONBASEMONEYS, COMMISSIONPCTMONEYS, TOTALCOMMISSIONMONEYS, CARNUM, GUSETNUM, EXCHANNO);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}'", EXCHANACOOUNT, PURGROUPSTARTDATES, GROUPSTARTDATES, PURGROUPENDDATES, GROUPENDDATES, STATUS);
                sbSql.AppendFormat(" )");
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

        public void UPDATEGROUPSALES(string ID, string CARNO, string CARNAME, string CARKIND, string GROUPKIND, string ISEXCHANGE, string CARNUM, string GUSETNUM, string EXCHANNO, string EXCHANACOOUNT,string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(" SET [CARNO]='{0}',[CARNAME]='{1}',[CARKIND]='{2}',[GROUPKIND]='{3}',[ISEXCHANGE]='{4}',[CARNUM]='{5}'", CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, CARNUM);
                sbSql.AppendFormat(" ,[GUSETNUM]='{0}',[EXCHANNO]='{1}',[EXCHANACOOUNT]='{2}',STATUS='{3}'", GUSETNUM, EXCHANNO, EXCHANACOOUNT, STATUS);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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

        public void UPDATEGROUPSALESCOMPELETE(string ID,string GROUPENDDATES, string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(" SET [GROUPENDDATES]='{0}',STATUS='{1}'", GROUPENDDATES, STATUS);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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
        public void SETMONEYS()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    //判断
                    if (dr.Cells[20].Value.ToString().Trim().Equals("預約接團"))
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
                        SPECIALMNUMS = FINDSPECIALMNUMS(ACCOUNT, STARTDATES, STARTTIMES);
                        SPECIALNUMSMONEYS = FINDSPECIALNUMSMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                        SPECIALMONEYS = FINDSPECIALMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                        SALESMMONEYS = FINDSALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);

                        //兌換券金額條件判斷
                        EXCHANGESALESMMONEYS = FINDEXCHANGESALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);

                        if(EXCHANGESALESMMONEYS>0)
                        {
                            int CARNUM = Convert.ToInt32(dr.Cells["車數"].Value.ToString().Trim());
                            EXCHANGEMONEYS = FINDEXCHANGEMONEYS();
                            EXCHANGETOTALMONEYS = EXCHANGEMONEYS * CARNUM;
                            //EXCHANGESALESMMONEYS = FINDEXCHANGESALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                            COMMISSIONBASEMONEYS = 0;

                            if (SALESMMONEYS > EXCHANGETOTALMONEYS)
                            {
                                SALESMMONEYS = SALESMMONEYS - EXCHANGETOTALMONEYS;
                            }
                        }
                        else
                        {
                            EXCHANGEMONEYS = 0;
                            EXCHANGETOTALMONEYS = 0;
                            EXCHANGESALESMMONEYS = 0;
                            COMMISSIONBASEMONEYS = FINDBASEMONEYS(CARKIND);
                        }

                      

                        SALESMMONEYS = SALESMMONEYS - SPECIALMONEYS;
                        COMMISSIONPCT = FINDCOMMISSIONPCT(CARKIND, SALESMMONEYS);
                        COMMISSIONPCTMONEYS = Convert.ToInt32(COMMISSIONPCT * SALESMMONEYS);
                        GUSETNUM = FINDGUSETNUM(ACCOUNT, STARTDATES, STARTTIMES);
                        TOTALCOMMISSIONMONEYS = Convert.ToInt32(SPECIALNUMSMONEYS + COMMISSIONBASEMONEYS + (COMMISSIONPCT * (SALESMMONEYS)));

                        UPDATEGROUPPRODUCT(ID, EXCHANGEMONEYS.ToString(), EXCHANGETOTALMONEYS.ToString(), EXCHANGESALESMMONEYS.ToString(), SALESMMONEYS.ToString(), SPECIALMNUMS.ToString(), SPECIALNUMSMONEYS.ToString(), COMMISSIONBASEMONEYS.ToString(), COMMISSIONPCT.ToString(), COMMISSIONPCTMONEYS.ToString(), TOTALCOMMISSIONMONEYS.ToString(), GUSETNUM.ToString());
                        //DateTime dt2 = DateTime.Now;

                        //MessageBox.Show(dt1.ToString("HH:mm:ss")+"-"+ dt2.ToString("HH:mm:ss"));
                    }

                }
            }

        }

        public int FINDEXCHANGEMONEYS()
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
               
                sbSql.AppendFormat(@"  SELECT  CONVERT(INT,[EXCHANGEMONEYS],0) AS EXCHANGEMONEYS   FROM [TKMK].[dbo].[GROUPEXCHANGEMONEYS]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public int FINDSPECIALMNUMS(string TA009,string TA001, string TA005)
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

                sbSql.AppendFormat(@"  SELECT  CONVERT(INT,ISNULL(SUM(TB019),0),0) AS 'SPECIALMNUMS'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006");
                sbSql.AppendFormat(@"  AND TB010 IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT])");
                sbSql.AppendFormat(@"  AND TA009='{0}'", TA009);
                sbSql.AppendFormat(@"  AND TA001='{0}'", TA001);
                sbSql.AppendFormat(@"  AND TA005>='{0}'", TA005);
                sbSql.AppendFormat(@"  ");

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  CONVERT(INT,ISNULL(SUM(TB033),0),0) AS 'SPECIALMONEYS'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006");
                sbSql.AppendFormat(@"  AND TB010 IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT] WHERE [SPLITCAL]='Y') ");
                sbSql.AppendFormat(@"  AND TA009='{0}'", TA009);
                sbSql.AppendFormat(@"  AND TA001='{0}'", TA001);
                sbSql.AppendFormat(@"  AND TA005>='{0}'", TA005);
                sbSql.AppendFormat(@"  ");

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

        public int FINDSPECIALNUMSMONEYS(string TA009, string TA001, string TA005)
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

                sbSql.AppendFormat(@"  
                                    SELECT CONVERT(INT,ISNULL(SUM(TB019/[GROUPPRODUCT].NUM*[GROUPPRODUCT].MONEYS),0),0) AS 'SPECIALNUMSMONEYS' 
                                    FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)
                                    LEFT JOIN [TKMK].[dbo].[GROUPPRODUCT] ON [GROUPPRODUCT].ID=TB010
                                    WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006
                                    AND TB010 IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT] WHERE [SPLITCAL]='Y')
                                    AND TA009='{0}'
                                    AND TA001='{1}'
                                    AND TA005>='{2}'
                                    ", TA009, TA001, TA005);
               

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["SPECIALNUMSMONEYS"].ToString());

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(INT,ISNULL(SUM(TB033),0),0) AS 'SALESMMONEYS'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTB WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006");                
                sbSql.AppendFormat(@"  AND TA009='{0}'", TA009);
                sbSql.AppendFormat(@"  AND TA001='{0}'", TA001);
                sbSql.AppendFormat(@"  AND TA005>='{0}'", TA005);
                sbSql.AppendFormat(@"  ");

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(INT,ISNULL(SUM(TA017),0)) AS EXCHANGESALESMMONEYS");
                sbSql.AppendFormat(@"  FROM [TK].dbo.POSTA WITH (NOLOCK),[TK].dbo.POSTC WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA001=TC001 AND TA002=TC002 AND TA003=TC003  AND TA006=TC006");
                sbSql.AppendFormat(@"  AND TC008='0009'");
                sbSql.AppendFormat(@"  AND TA009='{0}'", TA009);
                sbSql.AppendFormat(@"  AND TA001='{0}'", TA001);
                sbSql.AppendFormat(@"  AND TA005>='{0}'", TA005);
                sbSql.AppendFormat(@"  ");

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

        public int FINDBASEMONEYS(string NAME)
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

                sbSql.AppendFormat(@"  SELECT CONVERT(INT,[BASEMONEYS],0) AS 'BASEMONEYS' FROM [TKMK].[dbo].[GROUPBASE] WHERE [NAME]='{0}'", NAME);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public decimal FINDCOMMISSIONPCT(string CARKIND,int MONEYS)
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

               
                sbSql.AppendFormat(@"  SELECT [ID],[PCTMONEYS],[NAME],[PCT]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPPCT]");
                sbSql.AppendFormat(@"  WHERE [NAME]='{0}' AND ({1}-[PCTMONEYS])>=0", CARKIND, MONEYS);
                sbSql.AppendFormat(@"  ORDER BY ({0}-[PCTMONEYS])", MONEYS);
                sbSql.AppendFormat(@"  ");

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT COUNT(TA009) AS 'GUSETNUM'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.POSTA WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA009='{0}'", TA009);
                sbSql.AppendFormat(@"  AND TA001='{0}'", TA001);
                sbSql.AppendFormat(@"  AND TA005>='{0}'", TA005);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public void UPDATEGROUPPRODUCT(string ID, string EXCHANGEMONEYS, string EXCHANGETOTALMONEYS, string EXCHANGESALESMMONEYS, string SALESMMONEYS, string SPECIALMNUMS, string SPECIALMONEYS, string COMMISSIONBASEMONEYS,string COMMISSIONPCT, string COMMISSIONPCTMONEYS, string TOTALCOMMISSIONMONEYS,string GUSETNUM)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(" SET [EXCHANGEMONEYS]={0},[EXCHANGETOTALMONEYS]={1},[EXCHANGESALESMMONEYS]={2},[SALESMMONEYS]={3},[SPECIALMNUMS]={4},[SPECIALMONEYS]={5},[COMMISSIONBASEMONEYS]={6},[COMMISSIONPCT]={7},[COMMISSIONPCTMONEYS]={8},[TOTALCOMMISSIONMONEYS]={9},[GUSETNUM]={10}", EXCHANGEMONEYS, EXCHANGETOTALMONEYS, EXCHANGESALESMMONEYS, SALESMMONEYS, SPECIALMNUMS, SPECIALMONEYS, COMMISSIONBASEMONEYS, COMMISSIONPCT, COMMISSIONPCTMONEYS, TOTALCOMMISSIONMONEYS, GUSETNUM);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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
        }

        public void SETTEXT3()
        {
            textBox131.ReadOnly = false;
            textBox141.ReadOnly = false;
            textBox142.ReadOnly = false;
            textBox143.ReadOnly = false;

            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
          

        }

        public void SETTEXT4()
        {
            textBox131.ReadOnly = true;
            textBox141.ReadOnly = true;
            textBox142.ReadOnly = true;
            textBox143.ReadOnly = true;

            comboBox1.Enabled = false;
            comboBox2.Enabled = false;

        }
        public void SETTEXT5()
        {
            comboBox3.Enabled = true;
        }

        public void SETTEXT6()
        {
            checkBox1.Checked = false;
            comboBox3.Enabled = false;
        }


        public int SEARCHGROUPCAR(string CARNO)
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

                sbSql.AppendFormat(@"  SELECT [CARNO],[CARNAME],[CARKIND]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPCAR]");
                sbSql.AppendFormat(@"  WHERE [CARNO]='{0}'", CARNO);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public void SEARCHGROUPCARDETAIL(string CARNO)
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

                sbSql.AppendFormat(@"  SELECT [CARNO],[CARNAME],[CARKIND]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPCAR]");
                sbSql.AppendFormat(@"  WHERE [CARNO]='{0}'", CARNO);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    textBox141.Text=ds1.Tables["ds1"].Rows[0]["CARNAME"].ToString().Trim();
                    comboBox1.Text = ds1.Tables["ds1"].Rows[0]["CARKIND"].ToString().Trim();

                }
                else
                {
                    
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
        public void ADDGROUPCAR(string CARNO,string CARNAME,string CARKIND)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[GROUPCAR]");
                sbSql.AppendFormat(" ([CARNO],[CARNAME],[CARKIND])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}','{2}')", CARNO, CARNAME, CARKIND);
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

        public void UPDATEGROUPCAR(string CARNO, string CARNAME, string CARKIND)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" UPDATE [TKMK].[dbo].[GROUPCAR]");
                sbSql.AppendFormat(" SET [CARNAME]='{0}',[CARKIND]='{1}'", CARNAME, CARKIND);
                sbSql.AppendFormat(" WHERE [CARNO]='{0}'", CARNO);
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
        private void textBox131_TextChanged(object sender, EventArgs e)
        {
            SEARCHGROUPCARDETAIL(textBox131.Text.Trim());
        }


        public void SETNUMS(string GROUPSTARTDATES)
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

                sbSql.AppendFormat(@"  SELECT COUNT(CARNO) AS NUMS  ");
                sbSql.AppendFormat(@"  ,(SELECT SUM(GUSETNUM) FROM [TKMK].[dbo].[GROUPSALES] GP WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) ) AS GUSETNUMS");
                sbSql.AppendFormat(@"  ,(SELECT SUM(SALESMMONEYS) FROM [TKMK].[dbo].[GROUPSALES] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) ) AS SALESMMONEYS");
                sbSql.AppendFormat(@"  ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALES] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) AND STATUS='預約接團') AS CARNUM1");
                sbSql.AppendFormat(@"  ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALES] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) AND STATUS='取消預約') AS CARNUM2");
                sbSql.AppendFormat(@"  ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALES] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) AND STATUS='異常結案') AS CARNUM3");
                sbSql.AppendFormat(@"  ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALES] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) AND STATUS='完成接團') AS CARNUM4");
                sbSql.AppendFormat(@"  ,(SELECT COUNT(CARNO) FROM [TKMK].[dbo].[GROUPSALES] GP WITH(NOLOCK) WHERE CONVERT(NVARCHAR,GP.GROUPSTARTDATES,112)=CONVERT(NVARCHAR,[GROUPSALES].GROUPSTARTDATES,112) AND STATUS='預約接團') AS CARNUM5");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPSALES] WITH(NOLOCK)");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,GROUPSTARTDATES,112)='{0}'", GROUPSTARTDATES);
                sbSql.AppendFormat(@"  AND STATUS IN ('預約接團','完成接團')");
                sbSql.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR,GROUPSTARTDATES,112)");             
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    label11.Text = ds1.Tables["ds1"].Rows[0]["NUMS"].ToString().Trim();
                    label13.Text = ds1.Tables["ds1"].Rows[0]["GUSETNUMS"].ToString().Trim();
                    label15.Text = ds1.Tables["ds1"].Rows[0]["SALESMMONEYS"].ToString().Trim();
                    label17.Text = ds1.Tables["ds1"].Rows[0]["CARNUM1"].ToString().Trim();
                    label23.Text = ds1.Tables["ds1"].Rows[0]["CARNUM2"].ToString().Trim();
                    label20.Text = ds1.Tables["ds1"].Rows[0]["CARNUM3"].ToString().Trim();      
                    label24.Text = ds1.Tables["ds1"].Rows[0]["CARNUM4"].ToString().Trim();
                    label21.Text = ds1.Tables["ds1"].Rows[0]["CARNUM5"].ToString().Trim();

                }
                else
                {
                    label11.Text ="0";
                    label13.Text = "0";
                    label15.Text = "0";
                    label17.Text = "0";
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

        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

           

            Report report1 = new Report();
            if (comboBox5.Text.Equals("遊覽車對帳明細表"))
            {
                report1.Load(@"REPORT\遊覽車對帳明細表.frx");

                SQL = SETSQL();
            }
            else if (comboBox5.Text.Equals("多年期月份團務比較表"))
            {
                report1.Load(@"REPORT\多年期月份團務比較表.frx");

                SQL = SETSQL2();
            }

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();
            report1.SetParameterValue("P1", dateTimePicker4.Value.ToString("yyyy/MM/dd"));
            report1.SetParameterValue("P2", dateTimePicker5.Value.ToString("yyyy/MM/dd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {           
            StringBuilder SB = new StringBuilder();
           
            SB.AppendFormat(@" SELECT ");
            SB.AppendFormat(@" [GROUPSALES].[SERNO] AS '序號',CONVERT(NVARCHAR,[PURGROUPSTARTDATES],111) AS '日期',[CARNAME] AS '車名',[CARKIND] AS '車種',[CARNO] AS '車號',[CARNUM] AS '車數',[GROUPKIND] AS '團類',[GUSETNUM] AS '來客數',[EXCHANNO] AS '優惠券',[EXCHANACOOUNT] AS '優惠號',[ISEXCHANGE] AS '領兌'");
            SB.AppendFormat(@" ,[EXCHANGETOTALMONEYS] AS '兌換券金額',[EXCHANGESALESMMONEYS] AS '(兌)消費金額',[COMMISSIONBASEMONEYS] AS '茶水費',[SALESMMONEYS] AS '消費總額',[SPECIALMNUMS] AS '特賣組數',[SPECIALMONEYS] AS '特賣獎金',[COMMISSIONPCTMONEYS] AS '消費獎金',[TOTALCOMMISSIONMONEYS] AS '獎金合計',[STATUS] AS '狀態'");
            SB.AppendFormat(@" FROM [TKMK].[dbo].[GROUPSALES] WITH (NOLOCK) ");
            SB.AppendFormat(@" WHERE CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112)>='{0}' AND CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112)<='{1}'",dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@" AND [STATUS]='完成接團'");
            SB.AppendFormat(@" ORDER BY CONVERT(NVARCHAR,[PURGROUPSTARTDATES], 112),[SERNO]");
            SB.AppendFormat(@" ");

            return SB;

        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" SELECT SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 ) AS '年月'");
            SB.AppendFormat(@" ,(SELECT ISNULL(SUM(GS.[GUSETNUM]),0) FROM[TKMK].[dbo].[GROUPSALES] GS WITH (NOLOCK) WHERE CONVERT(NVARCHAR,GS.[PURGROUPSTARTDATES],112) LIKE SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 )+'%') AS '來客數'");
            SB.AppendFormat(@" ,(SELECT ISNULL(SUM(GS.[CARNUM]),0) FROM[TKMK].[dbo].[GROUPSALES] GS WITH (NOLOCK) WHERE CONVERT(NVARCHAR,GS.[PURGROUPSTARTDATES],112) LIKE SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 )+'%') AS '來車數'");
            SB.AppendFormat(@" ,(SELECT ISNULL(SUM(GS.[SALESMMONEYS]),0) FROM[TKMK].[dbo].[GROUPSALES] GS  WITH (NOLOCK) WHERE CONVERT(NVARCHAR,GS.[PURGROUPSTARTDATES],112) LIKE SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 )+'%') AS '團客總金額'");
            SB.AppendFormat(@" ,(SELECT SUM(ISNULL(TA017,0)) FROM [TK].dbo.POSTA WITH (NOLOCK) WHERE  TA002='106701' AND TA001 LIKE SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 )+'%') AS '消費總金額'");
            SB.AppendFormat(@" ,((SELECT SUM(ISNULL(TA017,0)) FROM [TK].dbo.POSTA WITH (NOLOCK) WHERE TA002='106701' AND TA001 LIKE SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 )+'%')-(SELECT ISNULL(SUM(GS.[SALESMMONEYS]),0) FROM[TKMK].[dbo].[GROUPSALES] GS WITH (NOLOCK) WHERE CONVERT(NVARCHAR,GS.[PURGROUPSTARTDATES],112) LIKE SUBSTRING(CONVERT(NVARCHAR,[GROUPSALES].[PURGROUPSTARTDATES],112),1,6 )+'%')) AS '散客總金額'");
            SB.AppendFormat(@" FROM [TKMK].[dbo].[GROUPSALES] WITH (NOLOCK)");
            SB.AppendFormat(@" WHERE CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112)>='{0}' AND CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112)<='{1}'", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@" AND [STATUS]='完成接團'");
            SB.AppendFormat(@" GROUP BY SUBSTRING(CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112),1,6 )");
            SB.AppendFormat(@" ORDER BY SUBSTRING(CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112),1,6 )");
            SB.AppendFormat(@"  ");
            SB.AppendFormat(@" ");
            SB.AppendFormat(@" ");



            return SB;

        }

        public void UPDATETKWSCMI(string MI001,string NAME)
        {
            try
            {
               
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(" UPDATE [TK].[dbo].[WSCMI] SET [MI002]='{0}' WHERE MI001='{1}'",NAME,MI001);
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

        public int FINDSEARCHGROUPSALESNOTFINISH(string CREATEDATES)
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

                sbSql.AppendFormat(@"  SELECT COUNT([CARNO]) AS NUMS ");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(@"  WHERE [STATUS]='預約接團' AND CONVERT(nvarchar,[CREATEDATES],112)='{0}' ", CREATEDATES);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void button5_Click(object sender, EventArgs e)
        {
            STATUSCONTROLLER = "ADD";

            SETTEXT1();
            comboBox3load();



        }
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));

            SETMONEYS();

            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));

            SETNUMS(dateTimePicker1.Value.ToString("yyyyMMdd"));

            label29.Text = "";
            label29.Text = "更新時間" + dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");
        }
        private void button9_Click(object sender, EventArgs e)
        {
            if(STATUSCONTROLLER.Equals("ADD"))
            {
                string ID = Guid.NewGuid().ToString();
                string CREATEDATES = dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string SERNO = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
                string CARNO = textBox131.Text.Trim();
                string CARNAME = textBox141.Text.Trim();
                string CARKIND = comboBox1.Text.Trim();
                string GROUPKIND = comboBox2.Text.Trim();
                string ISEXCHANGE = "N";

                if (checkBox1.Checked == true)
                {
                    ISEXCHANGE = "是";
                }
                else
                {
                    ISEXCHANGE = "否";
                }

                string EXCHANGEMONEYS = "0";
                string EXCHANGETOTALMONEYS = "0";
                string EXCHANGESALESMMONEYS = "0";
                string SALESMMONEYS = "0";
                string SPECIALMNUMS = "0";
                string SPECIALMONEYS = "0";
                string COMMISSIONBASEMONEYS = "0";
                string COMMISSIONPCTMONEYS = "0";
                string TOTALCOMMISSIONMONEYS = "0";
                string CARNUM = textBox142.Text.Trim();
                string GUSETNUM = textBox143.Text.Trim();
                string EXCHANNO = textBox144.Text.Trim();
                string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0,7).ToString();
                string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");
                string GROUPENDDATES = "1911/1/1";
                string STATUS = "預約接團";

                if (!string.IsNullOrEmpty(SERNO) && !string.IsNullOrEmpty(CARNO) && !string.IsNullOrEmpty(EXCHANNO) && !string.IsNullOrEmpty(EXCHANACOOUNT))
                {
                    ADDGROUPSALES(
                    ID, CREATEDATES, SERNO, CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, EXCHANGEMONEYS, EXCHANGETOTALMONEYS
                    , EXCHANGESALESMMONEYS, SALESMMONEYS, SPECIALMNUMS, SPECIALMONEYS, COMMISSIONBASEMONEYS, COMMISSIONPCTMONEYS, TOTALCOMMISSIONMONEYS, CARNUM, GUSETNUM, EXCHANNO
                    , EXCHANACOOUNT, PURGROUPSTARTDATES, GROUPSTARTDATES, PURGROUPENDDATES, GROUPENDDATES, STATUS
                   );
                    
                    textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
                    SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                }
                else
                {
                    MessageBox.Show("團務資料少填");
                }

                if(!string.IsNullOrEmpty(CARNO)&& !string.IsNullOrEmpty(CARNAME) && !string.IsNullOrEmpty(CARKIND) )
                {
                    int ISCAR=SEARCHGROUPCAR(CARNO);

                    if(ISCAR==0)
                    {
                        ADDGROUPCAR(CARNO, CARNAME, CARKIND);
                    }
                    else if(ISCAR==1)
                    {
                        UPDATEGROUPCAR(CARNO, CARNAME, CARKIND);
                    }
                }

                UPDATETKWSCMI(EXCHANACOOUNT, EXCHANNO+' '+CARNAME);
            }
            else if(STATUSCONTROLLER.Equals("EDIT"))
            {
                if(!string.IsNullOrEmpty(ID))
                {
                    string CARNO = textBox131.Text.Trim();
                    string CARNAME = textBox141.Text.Trim();
                    string CARKIND = comboBox1.Text.Trim();
                    string GROUPKIND = comboBox2.Text.Trim();
                    string ISEXCHANGE = "N";

                    if (checkBox1.Checked == true)
                    {
                        ISEXCHANGE = "是";
                    }
                    else
                    {
                        ISEXCHANGE = "否";
                    }
                    string CARNUM = textBox142.Text.Trim();
                    string GUSETNUM = textBox143.Text.Trim();
                    string EXCHANNO = textBox144.Text.Trim();
                    string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString();
                    //string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");

                    if (!string.IsNullOrEmpty(ID) && !string.IsNullOrEmpty(CARNO) && !string.IsNullOrEmpty(EXCHANNO) && !string.IsNullOrEmpty(EXCHANACOOUNT))
                    {
                        UPDATEGROUPSALES(ID, CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, CARNUM, GUSETNUM, EXCHANNO, EXCHANACOOUNT, "預約接團");
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
        private void button1_Click_1(object sender, EventArgs e)
        {
            if(STATUS.Equals("預約接團"))
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

        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES2(dateTimePicker1.Value.ToString("yyyyMMdd"),comboBox4.Text.Trim());
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
                    string ISEXCHANGE = "N";

                    if (checkBox1.Checked == true)
                    {
                        ISEXCHANGE = "是";
                    }
                    else
                    {
                        ISEXCHANGE = "否";
                    }
                    string CARNUM = textBox142.Text.Trim();
                    string GUSETNUM = textBox143.Text.Trim();
                    string EXCHANNO = textBox144.Text.Trim();
                    string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString(); 
                    //string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");

                    UPDATEGROUPSALES(ID, CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, CARNUM, GUSETNUM, EXCHANNO, EXCHANACOOUNT, "取消預約");
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
                    string ISEXCHANGE = "N";

                    if (checkBox1.Checked == true)
                    {
                        ISEXCHANGE = "是";
                    }
                    else
                    {
                        ISEXCHANGE = "否";
                    }
                    string CARNUM = textBox142.Text.Trim();
                    string GUSETNUM = textBox143.Text.Trim();
                    string EXCHANNO = textBox144.Text.Trim();
                    string EXCHANACOOUNT = comboBox3.Text.Trim().Substring(0, 7).ToString();
                    //string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");

                    UPDATEGROUPSALES(ID, CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, CARNUM, GUSETNUM, EXCHANNO, EXCHANACOOUNT, "異常結案");
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
                    string GROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");

                    UPDATEGROUPSALESCOMPELETE(ID, GROUPENDDATES, "完成接團");
                }

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                MessageBox.Show("不是預約接團，不能修改");
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

       

        #endregion


    }
}
