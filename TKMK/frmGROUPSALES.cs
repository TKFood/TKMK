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

        string STATUSCONTROLLER = null;
       
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
            timer1.Interval = 1000 * 10;
            timer1.Start();
        }

        #region FUNCTION

        private void timer1_Tick(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;

            textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
            comboBox3load();

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
            Sequel.AppendFormat(@"SELECT MI001,MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%'  AND MI001 NOT IN (SELECT [EXCHANACOOUNT] FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)=CONVERT(nvarchar,GETDATE(),112)) ORDER BY MI001 ");
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
            SEARCHWSCMI(comboBox3.Text.Trim());
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

                sbSql.AppendFormat(@"  SELECT MI001,MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%' AND MI001='{0}' ORDER BY MI001 ",MI001);
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

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [SERNO] AS '序號',[CARNO] AS '車號',[CARNAME] AS '車名',[CARKIND] AS '車種',[GROUPKIND]  AS '團類',[ISEXCHANGE] AS '兌換券',[EXCHANGEMONEYS] AS '領券額',[EXCHANGETOTALMONEYS] AS '券總額',[EXCHANGESALESMMONEYS] AS '券消費',[SALESMMONEYS] AS '消費總額'");
                sbSql.AppendFormat(@"  ,[SPECIALMNUMS] AS '特賣數',[SPECIALMONEYS] AS '特賣獎金',[COMMISSIONBASEMONEYS] AS '茶水費',[COMMISSIONPCTMONEYS] AS '消費獎金',[TOTALCOMMISSIONMONEYS] AS '總獎金',[CARNUM] AS '車數',[GUSETNUM] AS '人數',[EXCHANNO] AS '優惠券名',[EXCHANACOOUNT] AS '優惠券帳號',CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間',CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'");
                sbSql.AppendFormat(@"  ,CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間',CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間',[STATUS] AS '狀態',[ID],[CREATEDATES]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(@"  WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'", CREATEDATES);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(nvarchar,[CREATEDATES],112),[SERNO]");
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

                //add ZWAREWHOUSEPURTH
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

        public void SETMONEYS()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    string ACCOUNT= dr.Cells["品號"].Value.ToString().Trim();
                }
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

            
           

        }
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));

            SETMONEYS();
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
                    ISEXCHANGE = "Y";
                }
                else
                {
                    ISEXCHANGE = "N";
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
                string EXCHANACOOUNT = comboBox3.Text.Trim();
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
            }
            else
            {

            }


            STATUSCONTROLLER = null;
        }

        #endregion


    }
}
