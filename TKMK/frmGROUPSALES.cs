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
        string ID = null;
        string ACCOUNT = null;
        string ISEXCHANGE = null;
        string CARKIND = null;
        string GROUPSTARTDATES = null;
        string STARTDATES = null;
        string STARTTIMES = null;
        int SPECIALMNUMS = 0;
        int SPECIALMONEYS = 0;
        int EXCHANGEMONEYS = 0;
        int EXCHANGETOTALMONEYS = 0;
        int EXCHANGESALESMMONEYS = 0;
        int COMMISSIONBASEMONEYS = 0;
        int SALESMMONEYS = 0;
        decimal COMMISSIONPCT = 0;
        int COMMISSIONPCTMONEYS = 0;
        int TOTALCOMMISSIONMONEYS = 0;
        int GUSETNUM = 0;

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
            //dateTimePicker1.Value = DateTime.Now;
            //dateTimePicker2.Value = DateTime.Now;
            //dateTimePicker3.Value = DateTime.Now;

            //textBox121.Text = FINDSERNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
            //comboBox3load();

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
            Sequel.AppendFormat(@"SELECT MI001,MI002 FROM [TK].dbo.WSCMI WHERE MI001 LIKE '68%'  AND MI001 NOT IN (SELECT [EXCHANACOOUNT] FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}'  AND [STATUS]='預約接團' ) ORDER BY MI001 ",dateTimePicker1.Value.ToString("yyyyMMdd"));
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

        public void SEARCHGROUPSALES(string CREATEDATES,string STATUS)
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
                sbSql.AppendFormat(@"  ,[SPECIALMNUMS] AS '特賣數',[SPECIALMONEYS] AS '特賣獎金',[COMMISSIONBASEMONEYS] AS '茶水費',[COMMISSIONPCTMONEYS] AS '消費獎金',[TOTALCOMMISSIONMONEYS] AS '總獎金',[CARNUM] AS '車數',[GUSETNUM] AS '來客數',[EXCHANNO] AS '優惠券名',[EXCHANACOOUNT] AS '優惠券帳號',CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間',CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'");
                sbSql.AppendFormat(@"  ,CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間',CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間',[STATUS] AS '狀態',[COMMISSIONPCT] AS '抽佣比率',[ID],[CREATEDATES]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(@"  WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}' AND [STATUS]='{1}'", CREATEDATES, STATUS);
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();

                    textBox121.Text = row.Cells["序號"].Value.ToString();
                    textBox131.Text = row.Cells["車號"].Value.ToString();
                    textBox141.Text = row.Cells["車名"].Value.ToString();
                    textBox142.Text = row.Cells["車數"].Value.ToString();
                    textBox143.Text = row.Cells["來客數"].Value.ToString();
                    textBox144.Text = row.Cells["優惠券名"].Value.ToString();

                    comboBox1.Text = row.Cells["車種"].Value.ToString();
                    comboBox2.Text = row.Cells["團類"].Value.ToString();
                    comboBox3.Text = row.Cells["優惠券帳號"].Value.ToString();

                    if(row.Cells["兌換券"].Value.ToString().Equals("Y"))
                    {
                        checkBox1.Checked = true;
                    }
                    else if(row.Cells["兌換券"].Value.ToString().Equals("N"))
                    {
                        checkBox1.Checked = false;
                    }
                   
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

        public void UPDATEGROUPSALES(string ID, string CARNO, string CARNAME, string CARKIND, string GROUPKIND, string ISEXCHANGE, string CARNUM, string GUSETNUM, string EXCHANNO, string EXCHANACOOUNT)
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
                sbSql.AppendFormat(" ,[GUSETNUM]='{0}',[EXCHANNO]='{1}',[EXCHANACOOUNT]='{2}'", GUSETNUM, EXCHANNO, EXCHANACOOUNT);
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
                    //清空值
                    ID = null;
                    STATUSCONTROLLER = null;
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
                    ISEXCHANGE= dr.Cells["兌換券"].Value.ToString().Trim();
                    CARKIND= dr.Cells["車種"].Value.ToString().Trim();
                    GROUPSTARTDATES= dr.Cells["實際到達時間"].Value.ToString().Trim();
                    STARTDATES = GROUPSTARTDATES.Substring(0,10).Replace("-","").ToString();
                    STARTTIMES = GROUPSTARTDATES.Substring(11,8);

                    //DateTime dt1 = DateTime.Now;

                    //找出各項金額                   
                    SPECIALMNUMS = FINDSPECIALMNUMS(ACCOUNT, STARTDATES, STARTTIMES);
                    SPECIALMONEYS = FINDSPECIALMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                    SALESMMONEYS = FINDSALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);                  

                    //金額條件判斷
                    if (ISEXCHANGE.Equals("是"))
                    {
                        int CARNUM = Convert.ToInt32(dr.Cells["車數"].Value.ToString().Trim());
                        EXCHANGEMONEYS = FINDEXCHANGEMONEYS();
                        EXCHANGETOTALMONEYS = EXCHANGEMONEYS * CARNUM;
                        EXCHANGESALESMMONEYS = FINDEXCHANGESALESMMONEYS(ACCOUNT, STARTDATES, STARTTIMES);
                        COMMISSIONBASEMONEYS = 0;

                        SALESMMONEYS = SALESMMONEYS - EXCHANGETOTALMONEYS;

                    }
                    else if(ISEXCHANGE.Equals("否"))
                    {
                        EXCHANGEMONEYS = 0;
                        EXCHANGETOTALMONEYS = 0;
                        EXCHANGESALESMMONEYS = 0;
                        COMMISSIONBASEMONEYS = FINDBASEMONEYS(CARKIND);
                        
                    }

                    COMMISSIONPCT = FINDCOMMISSIONPCT(CARKIND, SALESMMONEYS);
                    COMMISSIONPCTMONEYS = Convert.ToInt32(COMMISSIONPCT * SALESMMONEYS);
                    GUSETNUM = FINDGUSETNUM(ACCOUNT, STARTDATES, STARTTIMES);
                    TOTALCOMMISSIONMONEYS = Convert.ToInt32(SPECIALMONEYS + COMMISSIONBASEMONEYS + (COMMISSIONPCT * (SALESMMONEYS - SPECIALMONEYS)));

                    UPDATEGROUPPRODUCT(ID, EXCHANGEMONEYS.ToString(), EXCHANGETOTALMONEYS.ToString(), EXCHANGESALESMMONEYS.ToString(), SALESMMONEYS.ToString(), SPECIALMNUMS.ToString(), SPECIALMONEYS.ToString(), COMMISSIONBASEMONEYS.ToString(), COMMISSIONPCT.ToString(), COMMISSIONPCTMONEYS.ToString(), TOTALCOMMISSIONMONEYS.ToString() , GUSETNUM.ToString());
                    //DateTime dt2 = DateTime.Now;

                    //MessageBox.Show(dt1.ToString("HH:mm:ss")+"-"+ dt2.ToString("HH:mm:ss"));

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

                sbSql.AppendFormat(@"  SELECT CONVERT(INT,[BASEMONEYS],0) AS 'BASEMONEYS' FROM [TKMK].[dbo].[GROUPBASE] WHERE [NAME]='大巴'");
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

           
        }

        public void SETTEXT4()
        {
            textBox131.ReadOnly = true;
            textBox141.ReadOnly = true;
            textBox142.ReadOnly = true;
            textBox143.ReadOnly = true;
         
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
        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void button5_Click(object sender, EventArgs e)
        {
            STATUSCONTROLLER = "ADD";

            SETTEXT1();



        }
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"),"預約接團");

            SETMONEYS();

            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"), "預約接團");
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
                    SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"), "預約接團");
                }
                else
                {
                    MessageBox.Show("團務資料少填");
                }
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
                    string EXCHANACOOUNT = comboBox3.Text.Trim();
                    //string PURGROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string GROUPSTARTDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss");
                    //string PURGROUPENDDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss");

                    UPDATEGROUPSALES(ID, CARNO, CARNAME, CARKIND, GROUPKIND, ISEXCHANGE, CARNUM, GUSETNUM, EXCHANNO, EXCHANACOOUNT);
                }
                
            }

            SETTEXT2();
            SETTEXT4();
            SETTEXT6();
            STATUSCONTROLLER = null;

            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"), "預約接團");
        }
        private void button10_Click(object sender, EventArgs e)
        {
            SETTEXT2();
            SETTEXT4();
            SETTEXT6();
            STATUSCONTROLLER = null;

            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"), "預約接團");
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            SETTEXT3();
            STATUSCONTROLLER = "EDIT";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            comboBox3load();
            SETTEXT5();
            STATUSCONTROLLER = "EDIT";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox4.Text.Trim());
        }

        #endregion


    }
}
