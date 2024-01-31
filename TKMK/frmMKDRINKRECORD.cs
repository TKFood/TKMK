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
        SqlDataAdapter adapter10 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder10 = new SqlCommandBuilder();
        SqlDataAdapter adapter11 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder11 = new SqlCommandBuilder();
        SqlDataAdapter adapter12 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder12 = new SqlCommandBuilder();
        SqlDataAdapter adapter13 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder13 = new SqlCommandBuilder();
        SqlDataAdapter adapter14 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder14 = new SqlCommandBuilder();
        SqlDataAdapter adapter15 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder15 = new SqlCommandBuilder();
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
        DataSet ds10 = new DataSet();
        DataSet ds11 = new DataSet();
        DataSet ds12 = new DataSet();
        DataSet ds13 = new DataSet();
        DataSet ds14 = new DataSet();
        DataSet ds15 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = "EDIT";
        string BUYNO;
        string OLDBUYNO;
        string CHECKDELBOMMD = "N";
        string CHECKDELINVTA = "N";

        string SID;
        string TD001;
        string TD002;
        string TD003;
        string TD004;
        string TD005;
        string TD007;
        string TD011;
        string BOMTDUDF01;

        string SID2;
        string TA001;
        string TA002;
        string TA003;
        string TA004;
        string TA005;
        string TA011;
        string TA029;
        string TB012;
        string INVTAUDF01;




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

        public class INVTADATA
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

            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA008;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA023;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA032;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA036;
            public string TA037;
            public string TA038;
            public string TA039;
            public string TA040;
            public string TA041;
            public string TA042;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string TA047;
            public string TA048;
            public string TA049;
            public string TA050;
            public string TA051;
            public string TA052;
            public string TA053;
            public string TA054;
            public string TA055;
            public string TA056;
            public string TA057;
            public string TA058;
            public string TA059;
            public string TA060;
            public string TA061;
            public string TA062;
            public string TA063;
            public string TA064;
            public string TA065;
            public string TA066;
            public string TA067;
            public string TA068;
            public string TA200;
           
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



        public frmMKDRINKRECORD()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE (ME002 NOT LIKE '%停用%' AND  ME001 NOT LIKE '118%') ORDER BY ME001 ");
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
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
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

            DRINKID.Text = comboBox2.SelectedValue.ToString();


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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE (ME002 NOT LIKE '%停用%' AND  ME001 NOT LIKE '118%') ORDER BY ME001 ");
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
            Sequel.AppendFormat(@"SELECT [ID],[DRINKNAME] FROM [TKMK].[dbo].[DRINKNAME] WHERE [USED]='Y' ORDER BY [ID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("DRINKNAME", typeof(string));

            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "DRINKNAME";
            sqlConn.Close();

            textBox23.Text = comboBox4.SelectedValue.ToString();


        }

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
            Sequel.AppendFormat(@"
                                SELECT 
                                 [KINDS]
                                ,[PARASNAMES]
                                ,[DVALUES]
                                FROM [TKMK].[dbo].[TBZPARAS]
                                WHERE  [KINDS]='frmMKDRINKRECORD'
                                ORDER BY [DVALUES] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARASNAMES", typeof(string));

            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "PARASNAMES";
            comboBox5.DisplayMember = "PARASNAMES";
            sqlConn.Close();

            DRINKID.Text = comboBox2.SelectedValue.ToString();


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            if(!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString())&&!comboBox1.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                textBox1.Text = comboBox1.SelectedValue.ToString();
            }
            
        }
        public void Search(string SDAYS,string EDAYS)
        {
            ds.Clear();

           
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
                                    CONVERT(NVARCHAR,[DATES],111) AS '日期' 
                                    ,[MV001] AS '工號' 
                                    ,[CARDNO] AS '卡號' 
                                    ,[NAMES] AS '姓名' 
                                    ,[DEP] AS '部門' 
                                    ,[DEPNAME] AS '部門名' 
                                    ,[DRINK] AS '飲品' 
                                    ,[OTHERS] AS '其他'
                                    ,[CUP] AS '數量' 
                                    ,[REASON] AS '原因' 
                                    ,[DRINKID] AS '品號' 
                                    ,[SIGN] AS '簽名' 
                                    ,[ID]
                                    FROM [TKMK].[dbo].[MKDRINKRECORD]
                                    WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}'  AND CONVERT(NVARCHAR,[DATES],112)<='{1}'  
                                    ORDER BY CONVERT(NVARCHAR,[DATES],111)
                                    ", SDAYS, EDAYS);

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


                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[DRINKID] AS '品號' ,[SIGN] AS '簽名' ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}'  ", dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"));
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

        public void Search4()
        {
            ds.Clear();


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


                sbSql.AppendFormat(@"  SELECT [DEP] AS '代號',[DEPNAME] AS '名稱'");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}'",dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  GROUP BY [DEP],[DEPNAME]");
                sbSql.AppendFormat(@"  ");

                adapter13 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder13 = new SqlCommandBuilder(adapter13);
                sqlConn.Open();
                ds13.Clear();
                adapter13.Fill(ds13, "ds13");
                sqlConn.Close();


                if (ds13.Tables["ds13"].Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds13.Tables["ds13"].Rows.Count >= 1)
                    {
                        dataGridView6.DataSource = ds13.Tables["ds13"];
                        dataGridView6.AutoResizeColumns();


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

        public void Search5()
        {
            ds.Clear();


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



                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[DRINKID] AS '品號' ,[SIGN] AS '簽名' ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[MKDRINKRECORD]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}'  ", dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND [DEP]='{0}'",textBox23.Text);
                sbSql.AppendFormat(@"  AND [ID] NOT IN  (SELECT UDF01 FROM [TK].dbo.INVTB WHERE ISNULL(UDF01,'')<>''  AND TB018 NOT IN ('V'))");
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],111)");
                sbSql.AppendFormat(@"   ");

                adapter14 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder14 = new SqlCommandBuilder(adapter14);
                sqlConn.Open();
                ds14.Clear();
                adapter14.Fill(ds14, "ds14");
                sqlConn.Close();


                if (ds14.Tables["ds14"].Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;
                }
                else
                {
                    if (ds14.Tables["ds14"].Rows.Count >= 1)
                    {
                        dataGridView7.DataSource = ds14.Tables["ds14"];
                        dataGridView7.AutoResizeColumns();


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

        public void Search6()
        {
            ds.Clear();


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

                sbSql.AppendFormat(@"  SELECT [DEP] AS  '部門', CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS  '費用單別',[TA002] AS  '費用單號'");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[INVTARESLUT] ");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}'  ", dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND [DEP]='{0}'", textBox23.Text);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],111)");
                sbSql.AppendFormat(@"  ");

                adapter15 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder15 = new SqlCommandBuilder(adapter15);
                sqlConn.Open();
                ds15.Clear();
                adapter15.Fill(ds15, "ds15");
                sqlConn.Close();


                if (ds15.Tables["ds15"].Rows.Count == 0)
                {
                    dataGridView8.DataSource = null;
                }
                else
                {
                    if (ds15.Tables["ds15"].Rows.Count >= 1)
                    {
                        dataGridView8.DataSource = ds15.Tables["ds15"];
                        dataGridView8.AutoResizeColumns();


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

        public void Search7(string SDAYS)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
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
                                   SELECT 
                                    CONVERT(NVARCHAR,[DATES],111) AS '日期' 
                                    ,[MV001] AS '工號' 
                                    ,[CARDNO] AS '卡號' 
                                    ,[NAMES] AS '姓名' 
                                    ,[DEP] AS '部門' 
                                    ,[DEPNAME] AS '部門名' 
                                    ,[DRINK] AS '飲品' 
                                    ,[OTHERS] AS '其他'
                                    ,[CUP] AS '數量' 
                                    ,[REASON] AS '原因' 
                                    ,[DRINKID] AS '品號' 
                                    ,[SIGN] AS '簽名' 
                                    ,[ID]
                                    FROM [TKMK].[dbo].[MKDRINKRECORD]
                                    WHERE CONVERT(NVARCHAR,[DATES],112)='{0}'  
                                    ORDER BY CONVERT(NVARCHAR,[DATES],111)
                                    ", SDAYS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView9.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView9.DataSource = ds.Tables["TEMPds1"];
                        dataGridView9.AutoResizeColumns();


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
                    textBox34.Text = row.Cells["工號"].Value.ToString();
                    textBox35.Text = row.Cells["姓名"].Value.ToString();
                    textBox36.Text = row.Cells["卡號"].Value.ToString();

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBoxID.Text = null;
                    textBox34.Text = null;
                    textBox35.Text = null;
                    textBox36.Text = null;

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
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = true;
        }
        public void UPDATE(
                            string DATES
                            , string MV001
                            , string CARDNO
                            , string NAMES
                            , string DEP
                            , string DEPNAME
                            , string DRINK
                            , string OTHERS
                            , string CUP
                            , string REASON
                            , string DRINKID
                            , string ID
            )
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();             
              
                sbSql.AppendFormat(@" 
                                    UPDATE [TKMK].[dbo].[MKDRINKRECORD]
                                    SET [DATES]='{1}'
                                    ,[MV001]='{2}'
                                    ,[CARDNO]='{3}'
                                    ,[NAMES]='{4}'
                                    ,[DEP]='{5}'
                                    ,[DEPNAME]='{6}'
                                    ,[DRINK]='{7}'
                                    ,[OTHERS]='{8}'
                                    ,[CUP]='{9}'
                                    ,[REASON]='{10}'
                                    ,[DRINKID]='{11}'
                                    WHERE [ID]='{0}'

                                    "
                                    , ID
                                    , DATES
                                    , MV001
                                    , CARDNO
                                    , NAMES
                                    , DEP
                                    , DEPNAME
                                    , DRINK
                                    , OTHERS
                                    , CUP
                                    , REASON
                                    , DRINKID
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

        public void ADD()
        {
            if(!string.IsNullOrEmpty(textBox1.Text.ToString()))
            {
                try
                {

                    //add ZWAREWHOUSEPURTH
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[MKDRINKRECORD]");
                    sbSql.AppendFormat(" ([DATES],[DEP],[DEPNAME],[DRINK],[OTHERS],[CUP],[REASON],[SIGN],[DRINKID])");
                    sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", dateTimePicker3.Value.ToString("yyyyMMdd"), textBox1.Text, comboBox1.Text, comboBox2.Text, textBox2.Text, textBox3.Text, textBox4.Text, null, comboBox2.SelectedValue.ToString());
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
        public void DEL()
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

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

                       
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


            
            FASTSQL.AppendFormat(@"  
                                SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期' ,[DEP] AS '部門' ,[DEPNAME] AS '部門名' ,[DRINK] AS '飲品' ,[OTHERS] AS '其他' ,[CUP] AS '數量' ,[REASON] AS '原因' ,[SIGN] AS '簽名' ,[ID]
                                FROM [TKMK].[dbo].[MKDRINKRECORD]
                                WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' 
                                  ORDER BY [NAMES]
                                ", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));

            return FASTSQL.ToString();
        }


        public void SETFASTREPORT2()
        {

            string SQL;
            Report report2 = new Report();
            report2.Load(@"REPORT\飲品記錄表加總.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report2.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            
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
            DRINKID.Text = comboBox2.SelectedValue.ToString();
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
                    BOMTDUDF01 = row.Cells["ID"].Value.ToString();

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
                    BOMTDUDF01 = null;
                }
            }
        }


        public void ADDBOMTDRESLUT()
        {
            if(!string.IsNullOrEmpty(SID) && !string.IsNullOrEmpty(TD001) &&!string.IsNullOrEmpty(TD002))
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

        public void ADDINVTARESLUT(string DEP,string DATES,string TA001,string TA002)
        {
            if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
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

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    //TD001 = "A421";
                    //TD002 = "20190930003";

                    sbSql.AppendFormat(" INSERT INTO [TKMK].[dbo].[INVTARESLUT]");
                    sbSql.AppendFormat(" ([DEP],[DATES],[TA001],[TA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')", DEP, DATES, TA001, TA002);
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
        public string GETMAXTD002(string TD001,string TD003)
        {
            string TD002;

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
                        TD002 = SETTD002(ds6.Tables["ds6"].Rows[0]["TD002"].ToString(), TD003);
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

        public string SETTD002(string TD002, string TD003)
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


        public string GETMAXTA002(string TA001,string TA003)
        {
            string TA002;

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
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[INVTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, TA003.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter10 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder10 = new SqlCommandBuilder(adapter10);
                sqlConn.Open();
                ds10.Clear();
                adapter10.Fill(ds10, "ds10");
                sqlConn.Close();


                if (ds10.Tables["ds10"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds10.Tables["ds10"].Rows.Count >= 1)
                    {
                        TA002 = SETTA002(ds10.Tables["ds10"].Rows[0]["TA002"].ToString(),TA003);
                        return TA002;

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




        public string SETTA002(string TA002,string TA003)
        {
            if (TA002.Equals("00000000000"))
            {
                return TA003.ToString() + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return TA003.ToString() + temp.ToString();
            }

            return null;
        }


        public void CHECKBOMTDRESLUT()
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

        public void CHECKINVTARESLUT()
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
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [TA001] AS '費用單' ,[TA002]  AS '費用號',[SID] AS '來源' ");
                sbSql.AppendFormat(@" FROM [TKMK].[dbo].[INVTARESLUT] ");
                sbSql.AppendFormat(@" WHERE [SID]='{0}' ", textBoxID3.Text.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter11 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder11 = new SqlCommandBuilder(adapter11);
                sqlConn.Open();
                ds11.Clear();
                adapter11.Fill(ds11, "ds11");
                sqlConn.Close();


                if (ds11.Tables["ds11"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds11.Tables["ds11"].Rows.Count >= 1)
                    {
                        dataGridView5.DataSource = ds11.Tables["ds11"];
                        dataGridView5.AutoResizeColumns();


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
           

            if (!string.IsNullOrEmpty(TD001)&& !string.IsNullOrEmpty(TD002) && !string.IsNullOrEmpty(TD003) && !string.IsNullOrEmpty(TD004) && !string.IsNullOrEmpty(TD005) && !string.IsNullOrEmpty(TD007) )
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
                    sbSql.AppendFormat(" ,'{0}' [TE001],'{1}' [TE002],[BOMMD].MD002 [TE003],[BOMMD].MD003 [TE004],[BOMMD].MD004 [TE005],[BOMMD].MD005 [TE006],[INVMB].MB017 [TE007],ROUND({2}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) [TE008],CASE WHEN MB2.MB022='N' THEN NULL ELSE '庫存量或批號量不足 !' END [TE009],'{4}' [TE010]", TD001, TD002, TD007, null, 'N');
                    sbSql.AppendFormat(" ,'{0}' [TE011],'{1}' [TE012], CASE WHEN MB2.MB022='N' THEN NULL ELSE '********************' END [TE013],'{3}' [TE014],'{4}' [TE015],'{5}' [TE016],'{6}' [TE017],'{7}' [TE018],'{8}' [TE019],'{9}' [TE020]", '0', '0',null  , null, null, '0', '0', null, null, null);
                    sbSql.AppendFormat(" ,'{0}' [TE021],'{1}' [TE022],'{2}' [TE023],'{3}' [TE024],'{4}' [TE025],'{5}' [TE026],'{6}' [TE027],'{7}' [TE028],'{8}' [TE029]", '0', null, '0', null, '0', null, null, null, '0');
                    sbSql.AppendFormat(" ,'{0}' [UDF01],'{1}' [UDF02],'{2}' [UDF03],'{3}' [UDF04],'{4}' [UDF05],'{5}' [UDF06],'{6}' [UDF07],'{7}' [UDF08],'{8}' [UDF09],'{9}' [UDF10]", null, null, null, null, null, '0', '0', '0', '0', '0');
                    sbSql.AppendFormat(" FROM [TKMK].[dbo].[MKDRINKRECORD],[TK].dbo.[INVMB],[TK].dbo.[BOMMD]");
                    sbSql.AppendFormat(" LEFT JOIN [TK].dbo.[INVMB] MB2 ON [BOMMD].MD003=MB2.MB001");
                    sbSql.AppendFormat(" WHERE [DRINKID]=[INVMB].MB001 AND [BOMMD].MD001=[INVMB].MB001 ");
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
            DataTable dt = SEARCH_DRINKCREATOR();

            BOMTDDATA BOMTD = new BOMTDDATA();
            BOMTD.COMPANY = "TK";
            BOMTD.CREATOR = dt.Rows[0]["CREATOR"].ToString();
            BOMTD.USR_GROUP = dt.Rows[0]["USR_GROUP"].ToString();
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            BOMTD.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            BOMTD.MODIFIER = dt.Rows[0]["CREATOR"].ToString();
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
            BOMTD.DataGroup = dt.Rows[0]["USR_GROUP"].ToString();
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
            BOMTD.TD023 = "";
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
            BOMTD.UDF01 = BOMTDUDF01;
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

        public void ADDINVTAB()
        {
            INVTADATA INVTA = new INVTADATA();
            INVTA = SETINVTA();

            if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) )
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

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[INVTA]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007],[TA008],[TA009],[TA010]");
                    sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018],[TA019],[TA020]");
                    sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025],[TA026],[TA027],[TA028],[TA029],[TA030]");
                    sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035],[TA036],[TA037],[TA038],[TA039],[TA040]");
                    sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045],[TA046],[TA047],[TA048],[TA049],[TA050]");
                    sbSql.AppendFormat(" ,[TA051],[TA052],[TA053],[TA054],[TA055],[TA056],[TA057],[TA058],[TA059],[TA060]");
                    sbSql.AppendFormat(" ,[TA061],[TA062],[TA063],[TA064],[TA065],[TA066],[TA067],[TA068],[TA200]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");                
                    sbSql.AppendFormat(" SELECT ");
                    sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", INVTA.COMPANY, INVTA.CREATOR, INVTA.USR_GROUP, INVTA.CREATE_DATE, INVTA.MODIFIER, INVTA.MODI_DATE, INVTA.FLAG, INVTA.CREATE_TIME, INVTA.MODI_TIME, INVTA.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],'{4}' [sync_count],'{5}' [DataUser],'{6}' [DataGroup]", INVTA.TRANS_NAME, INVTA.sync_date, INVTA.sync_time, INVTA.sync_mark, INVTA.sync_count, INVTA.DataUser, INVTA.DataGroup);
                    sbSql.AppendFormat(" ,'{0}' [TA001],'{1}' [TA002],'{2}' [TA003],'{3}' [TA004],'{4}' [TA005],'{5}' [TA006],{6} [TA007],'{7}' [TA008],'{8}' [TA009],{9} [TA010]",INVTA.TA001, INVTA.TA002, INVTA.TA003, INVTA.TA004, INVTA.TA005, INVTA.TA006, INVTA.TA007, INVTA.TA008, INVTA.TA009, INVTA.TA010);
                    sbSql.AppendFormat(" ,[CUP] [TA011],ISNULL((SELECT TOP 1 [LB010] FROM [TK].dbo.[INVLB] WHERE  LB001=[DRINKID] AND [LB010]>0 ORDER BY LB002 DESC)*[CUP],0)*[CUP] [TA012],'{0}' [TA013],'{1}' [TA014],'{2}' [TA015],{3} [TA016],'{4}' [TA017],'{5}' [TA018],{6} [TA019],'{7}' [TA020]", INVTA.TA013, INVTA.TA014, INVTA.TA015, INVTA.TA016, INVTA.TA017, INVTA.TA018, INVTA.TA019, INVTA.TA020);
                    sbSql.AppendFormat(" ,'{0}' [TA021],'{1}' [TA022],'{2}' [TA023],'{3}' [TA024],'{4}' [TA025],'{5}' [TA026],'{6}' [TA027],'{7}' [TA028],'{8}' [TA029],'{9}' [TA030]", INVTA.TA021, INVTA.TA022, INVTA.TA023, INVTA.TA024, INVTA.TA025, INVTA.TA026, INVTA.TA027, INVTA.TA028, INVTA.TA029, INVTA.TA030);
                    sbSql.AppendFormat(" ,'{0}' [TA031],'{1}' [TA032],{2} [TA033],{3} [TA034],'{4}' [TA035],'{5}' [TA036],'{6}' [TA037],'{7}' [TA038],'{8}' [TA039],{9} [TA040]", INVTA.TA031, INVTA.TA032, INVTA.TA033, INVTA.TA034, INVTA.TA035, INVTA.TA036, INVTA.TA037, INVTA.TA038, INVTA.TA039, INVTA.TA040);
                    sbSql.AppendFormat(" ,{0} [TA041],{1} [TA042],'{2}' [TA043],'{3}' [TA044],'{4}' [TA045],'{5}' [TA046],'{6}' [TA047],'{7}' [TA048],{8} [TA049],{9} [TA050]", INVTA.TA041, INVTA.TA042, INVTA.TA043, INVTA.TA044, INVTA.TA045, INVTA.TA046, INVTA.TA047, INVTA.TA048, INVTA.TA049, INVTA.TA050);
                    sbSql.AppendFormat(" ,'{0}' [TA051],'{1}' [TA052],'{2}' [TA053],'{3}' [TA054],{4} [TA055],{5} [TA056],{6} [TA057],{7} [TA058],'{8}' [TA059],'{9}' [TA060]", INVTA.TA051, INVTA.TA052, INVTA.TA053, INVTA.TA054, INVTA.TA055, INVTA.TA056, INVTA.TA057, INVTA.TA058, INVTA.TA059, INVTA.TA060);
                    sbSql.AppendFormat(" ,'{0}' [TA061],'{1}' [TA062],'{2}' [TA063],'{3}' [TA064],'{4}' [TA065],'{5}' [TA066],'{6}' [TA067],'{7}' [TA068],'{8}' [TA200]", INVTA.TA061, INVTA.TA062, INVTA.TA063, INVTA.TA064, INVTA.TA065, INVTA.TA066, INVTA.TA067, INVTA.TA068, INVTA.TA200);
                    sbSql.AppendFormat(" ,'{0}' [UDF01],'{1}' [UDF02],'{2}' [UDF03],'{3}' [UDF04],'{4}' [UDF05],'{5}' [UDF06],'{6}' [UDF07],'{7}' [UDF08],'{8}' [UDF09],'{9}' [UDF10]", INVTA.UDF01, INVTA.UDF02, INVTA.UDF03, INVTA.UDF04, INVTA.UDF05, INVTA.UDF06, INVTA.UDF07, INVTA.UDF08, INVTA.UDF09, INVTA.UDF10);
                    sbSql.AppendFormat(" FROM [TKMK].[dbo].[MKDRINKRECORD],[TK].dbo.[INVMB]");
                    sbSql.AppendFormat(" WHERE [DRINKID]=MB001 ");
                    sbSql.AppendFormat(" AND [ID]='{0}'",SID2);             
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[INVTB]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007],[TB008],[TB009],[TB010]");
                    sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015],[TB016],[TB017],[TB018],[TB019],[TB020]");
                    sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025],[TB026],[TB027],[TB028],[TB029],[TB030]");
                    sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]");
                    sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]");
                    sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]");
                    sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]");
                    sbSql.AppendFormat(" ,[TB071],[TB072],[TB073]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" SELECT ");
                    sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", INVTA.COMPANY, INVTA.CREATOR, INVTA.USR_GROUP, INVTA.CREATE_DATE, INVTA.MODIFIER, INVTA.MODI_DATE, INVTA.FLAG, INVTA.CREATE_TIME, INVTA.MODI_TIME, INVTA.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],'{4}' [sync_count],'{5}' [DataUser],'{6}' [DataGroup]", INVTA.TRANS_NAME, INVTA.sync_date, INVTA.sync_time, INVTA.sync_mark, INVTA.sync_count, INVTA.DataUser, INVTA.DataGroup);
                    sbSql.AppendFormat(" ,'{0}' [TB001],'{1}' [TB002],RIGHT(REPLICATE('0',4) + CAST(ROW_NUMBER() OVER(ORDER BY [ID])  as NVARCHAR),4) [TB003],MB001 [TB004],MB002 [TB005],MB003 [TB006],CUP [TB007],MB004 [TB008],{2} [TB009],ISNULL((SELECT TOP 1 [LB010] FROM [TK].dbo.[INVLB] WHERE  LB001=[DRINKID] AND [LB010]>0 ORDER BY LB002 DESC)*[CUP],0)  [TB010]",TA001,TA002,0);
                    sbSql.AppendFormat(" ,ISNULL((SELECT TOP 1 [LB010] FROM [TK].dbo.[INVLB] WHERE  LB001=[DRINKID] AND [LB010]>0 ORDER BY LB002 DESC)*[CUP],0)*[CUP] [TB011],'{0}' [TB012],'{1}' [TB013],'{2}' [TB014],'{3}' [TB015],'{4}' [TB016],REASON [TB017],'{5}' [TB018],'{6}' [TB019],'{7}' [TB020]", TB012,null,null,null,null,"N",TA003,null);
                    sbSql.AppendFormat(" ,'{0}' [TB021],{1} [TB022],'{2}' [TB023],'{3}' [TB024],MB047 [TB025],MB047*CUP [TB026],'{4}' [TB027],'{5}' [TB028],'{6}' [TB029],{7} [TB030]", null,0,null,"N",null,null, null, 0);
                    sbSql.AppendFormat(" ,'{0}' [TB031],'{1}' [TB032],'{2}' [TB033],'{3}' [TB034],'{4}' [TB035],'{5}' [TB036],{6} [TB037],{7} [TB038],{8} [TB039],'{9}' [TB040]",0,null, null, null, null, null,0,0,0, null, null);
                    sbSql.AppendFormat(" ,'{0}' [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],{4} [TB045],'{5}' [TB046],{6} [TB047],'{7}' [TB048],'{8}' [TB049],{9} [TB050]", null, null, null, null,0, null,0, null, null,0);
                    sbSql.AppendFormat(" ,'{0}' [TB051],'{1}' [TB052],'{2}' [TB053],'{3}' [TB054],{4} [TB055],'{5}' [TB056],'{6}' [TB057],'{7}' [TB058],{8} [TB059],{9} [TB060]", null,"N", null, null,0, null, null, null,0,0);
                    sbSql.AppendFormat(" ,'{0}' [TB061],{1} [TB062],'{2}' [TB063],{3} [TB064],{4} [TB065],{5} [TB066],{6} [TB067],'{7}' [TB068],'{8}' [TB069],'{9}' [TB070]", null,0, null,0,0,0,0, null, null, null);
                    sbSql.AppendFormat(" ,'{0}' [TB071],'{1}' [TB072],'{2}' [TB073]", null, null, null);
                    sbSql.AppendFormat(" ,'{0}' [UDF01],'{1}' [UDF02],'{2}' [UDF03],'{3}' [UDF04],'{4}' [UDF05],{5} [UDF06],{6} [UDF07],{7} [UDF08],{8} [UDF09],{9} [UDF10]", null, null, null, null, null,0,0,0,0,0);
                    sbSql.AppendFormat(" FROM [TKMK].[dbo].[MKDRINKRECORD],[TK].dbo.[INVMB]");
                    sbSql.AppendFormat(" WHERE [DRINKID]=MB001");
                    sbSql.AppendFormat(" AND [ID]='{0}'", SID2);
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
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

        public void ADDINVTAB2()
        {
            INVTADATA INVTA = new INVTADATA();
            INVTA = SETINVTA();
            INVTA.TA004 = textBox23.Text;

            if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
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

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[INVTA]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007],[TA008],[TA009],[TA010]");
                    sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018],[TA019],[TA020]");
                    sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025],[TA026],[TA027],[TA028],[TA029],[TA030]");
                    sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035],[TA036],[TA037],[TA038],[TA039],[TA040]");
                    sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045],[TA046],[TA047],[TA048],[TA049],[TA050]");
                    sbSql.AppendFormat(" ,[TA051],[TA052],[TA053],[TA054],[TA055],[TA056],[TA057],[TA058],[TA059],[TA060]");
                    sbSql.AppendFormat(" ,[TA061],[TA062],[TA063],[TA064],[TA065],[TA066],[TA067],[TA068],[TA200]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES ");
                    sbSql.AppendFormat(" (");               
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.COMPANY, INVTA.CREATOR, INVTA.USR_GROUP, INVTA.CREATE_DATE, INVTA.MODIFIER, INVTA.MODI_DATE, INVTA.FLAG, INVTA.CREATE_TIME, INVTA.MODI_TIME, INVTA.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}','{1}' ,'{2}','{3}','{4}','{5}','{6}' ", INVTA.TRANS_NAME, INVTA.sync_date, INVTA.sync_time, INVTA.sync_mark, INVTA.sync_count, INVTA.DataUser, INVTA.DataGroup);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}',{6},'{7}','{8}',{9}", INVTA.TA001, INVTA.TA002, INVTA.TA003, INVTA.TA004, INVTA.TA005, INVTA.TA006, INVTA.TA007, INVTA.TA008, INVTA.TA009, INVTA.TA010);
                    sbSql.AppendFormat(" , {0},{1},'{2}','{3}','{4}',{5},'{6}','{7}',{8},'{9}'", INVTA.TA011, INVTA.TA012, INVTA.TA013, INVTA.TA014, INVTA.TA015, INVTA.TA016, INVTA.TA017, INVTA.TA018, INVTA.TA019, INVTA.TA020);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA021, INVTA.TA022, INVTA.TA023, INVTA.TA024, INVTA.TA025, INVTA.TA026, INVTA.TA027, INVTA.TA028, INVTA.TA029, INVTA.TA030);
                    sbSql.AppendFormat(" ,'{0}','{1}',{2},{3},'{4}','{5}','{6}','{7}','{8}',{9} ", INVTA.TA031, INVTA.TA032, INVTA.TA033, INVTA.TA034, INVTA.TA035, INVTA.TA036, INVTA.TA037, INVTA.TA038, INVTA.TA039, INVTA.TA040);
                    sbSql.AppendFormat(" , {0},{1},'{2}','{3}','{4}','{5}','{6}','{7}',{8},{9}", INVTA.TA041, INVTA.TA042, INVTA.TA043, INVTA.TA044, INVTA.TA045, INVTA.TA046, INVTA.TA047, INVTA.TA048, INVTA.TA049, INVTA.TA050);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}',{4},{5},{6},{7} ,'{8}','{9}'", INVTA.TA051, INVTA.TA052, INVTA.TA053, INVTA.TA054, INVTA.TA055, INVTA.TA056, INVTA.TA057, INVTA.TA058, INVTA.TA059, INVTA.TA060);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}' ,'{6}','{7}','{8}' ", INVTA.TA061, INVTA.TA062, INVTA.TA063, INVTA.TA064, INVTA.TA065, INVTA.TA066, INVTA.TA067, INVTA.TA068, INVTA.TA200);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}' ,'{8}','{9}'", INVTA.UDF01, INVTA.UDF02, INVTA.UDF03, INVTA.UDF04, INVTA.UDF05, INVTA.UDF06, INVTA.UDF07, INVTA.UDF08, INVTA.UDF09, INVTA.UDF10);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[INVTB]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007],[TB008],[TB009],[TB010]");
                    sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015],[TB016],[TB017],[TB018],[TB019],[TB020]");
                    sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025],[TB026],[TB027],[TB028],[TB029],[TB030]");
                    sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]");
                    sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]");
                    sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]");
                    sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]");
                    sbSql.AppendFormat(" ,[TB071],[TB072],[TB073]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" SELECT ");
                    sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", INVTA.COMPANY, INVTA.CREATOR, INVTA.USR_GROUP, INVTA.CREATE_DATE, INVTA.MODIFIER, INVTA.MODI_DATE, INVTA.FLAG, INVTA.CREATE_TIME, INVTA.MODI_TIME, INVTA.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],'{4}' [sync_count],'{5}' [DataUser],'{6}' [DataGroup]", INVTA.TRANS_NAME, INVTA.sync_date, INVTA.sync_time, INVTA.sync_mark, INVTA.sync_count, INVTA.DataUser, INVTA.DataGroup);
                    sbSql.AppendFormat(" ,'{0}' [TB001],'{1}' [TB002],RIGHT(REPLICATE('0',4) + CAST(ROW_NUMBER() OVER(ORDER BY [ID])  as NVARCHAR),4) [TB003],MB001 [TB004],MB002 [TB005],MB003 [TB006],CUP [TB007],MB004 [TB008],{2} [TB009],ISNULL((SELECT TOP 1 [LB010] FROM [TK].dbo.[INVLB] WHERE  LB001=[DRINKID] AND [LB010]>0 ORDER BY LB002 DESC)*[CUP],0)  [TB010]", TA001, TA002, 0);
                    sbSql.AppendFormat(" ,ISNULL((SELECT TOP 1 [LB010] FROM [TK].dbo.[INVLB] WHERE  LB001=[DRINKID] AND [LB010]>0 ORDER BY LB002 DESC)*[CUP],0)*[CUP] [TB011],'{0}' [TB012],'{1}' [TB013],'{2}' [TB014],'{3}' [TB015],'{4}' [TB016],REASON [TB017],'{5}' [TB018],'{6}' [TB019],'{7}' [TB020]", TB012, null, null, null, null, "N", TA003, null);
                    sbSql.AppendFormat(" ,'{0}' [TB021],{1} [TB022],'{2}' [TB023],'{3}' [TB024],MB047 [TB025],MB047*CUP [TB026],'{4}' [TB027],'{5}' [TB028],'{6}' [TB029],{7} [TB030]", null, 0, null, "N", null, null, null, 0);
                    sbSql.AppendFormat(" ,'{0}' [TB031],'{1}' [TB032],'{2}' [TB033],'{3}' [TB034],'{4}' [TB035],'{5}' [TB036],{6} [TB037],{7} [TB038],{8} [TB039],'{9}' [TB040]", 0, null, null, null, null, null, 0, 0, 0, null, null);
                    sbSql.AppendFormat(" ,'{0}' [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],{4} [TB045],'{5}' [TB046],{6} [TB047],'{7}' [TB048],'{8}' [TB049],{9} [TB050]", null, null, null, null, 0, null, 0, null, null, 0);
                    sbSql.AppendFormat(" ,'{0}' [TB051],'{1}' [TB052],'{2}' [TB053],'{3}' [TB054],{4} [TB055],'{5}' [TB056],'{6}' [TB057],'{7}' [TB058],{8} [TB059],{9} [TB060]", null, "N", null, null, 0, null, null, null, 0, 0);
                    sbSql.AppendFormat(" ,'{0}' [TB061],{1} [TB062],'{2}' [TB063],{3} [TB064],{4} [TB065],{5} [TB066],{6} [TB067],'{7}' [TB068],'{8}' [TB069],'{9}' [TB070]", null, 0, null, 0, 0, 0, 0, null, null, null);
                    sbSql.AppendFormat(" ,'{0}' [TB071],'{1}' [TB072],'{2}' [TB073]", null, null, null);
                    sbSql.AppendFormat(" ,[MKDRINKRECORD].[ID] [UDF01],'{1}' [UDF02],'{2}' [UDF03],'{3}' [UDF04],'{4}' [UDF05],{5} [UDF06],{6} [UDF07],{7} [UDF08],{8} [UDF09],{9} [UDF10]", null, null, null, null, null, 0, 0, 0, 0, 0);
                    sbSql.AppendFormat(" FROM [TKMK].[dbo].[MKDRINKRECORD],[TK].dbo.[INVMB]");
                    sbSql.AppendFormat(" WHERE [DRINKID]=MB001");
                    sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}'  ", dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(" AND [DEP]='{0}'", textBox23.Text);
                    sbSql.AppendFormat(" AND UDF01 NOT IN  (SELECT UDF01 FROM [TK].dbo.INVTB WHERE ISNULL(UDF01,'')<>'')  ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
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

        public INVTADATA SETINVTA()
        {
            DataTable dt = SEARCH_DRINKCREATOR();

            INVTADATA INVTA = new INVTADATA();
            INVTA.COMPANY = "TK";
            INVTA.CREATOR = dt.Rows[0]["CREATOR"].ToString();
            INVTA.USR_GROUP = dt.Rows[0]["USR_GROUP"].ToString();
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            INVTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            INVTA.MODIFIER = dt.Rows[0]["CREATOR"].ToString();
            INVTA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            INVTA.FLAG = "0";
            INVTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            INVTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            INVTA.TRANS_TYPE = "P001";
            INVTA.TRANS_NAME = "INVMI05";
            INVTA.sync_date = "";
            INVTA.sync_time = "";
            INVTA.sync_mark = "";
            INVTA.sync_count = "0";
            INVTA.DataUser = "";
            INVTA.DataGroup = dt.Rows[0]["USR_GROUP"].ToString();
            INVTA.TA001=TA001;
            INVTA.TA002=TA002;
            INVTA.TA003=TA003;
            INVTA.TA004=TA004;
            INVTA.TA005=TA005;
            INVTA.TA006="N";
            INVTA.TA007="0";
            INVTA.TA008="20";
            INVTA.TA009="11";
            INVTA.TA010="0";
            INVTA.TA011= "0";
            INVTA.TA012="0";
            INVTA.TA013="N";
            INVTA.TA014=TA003;
            INVTA.TA015="";
            INVTA.TA016="0";
            INVTA.TA017="N";
            INVTA.TA018="N";
            INVTA.TA019="0";
            INVTA.TA020="6";
            INVTA.TA021="";
            INVTA.TA022="";
            INVTA.TA023="";
            INVTA.TA024="";
            INVTA.TA025="";
            INVTA.TA026="";
            INVTA.TA027="";
            INVTA.TA028="";
            INVTA.TA029=TA029;
            INVTA.TA030="";
            INVTA.TA031="";
            INVTA.TA032="";
            INVTA.TA033 = "0";
            INVTA.TA034 = "0";
            INVTA.TA035 = "";
            INVTA.TA036 = "";
            INVTA.TA037 = "";
            INVTA.TA038 = "";
            INVTA.TA039 = "";
            INVTA.TA040 = "0";
            INVTA.TA041 = "0";
            INVTA.TA042 = "0";
            INVTA.TA043 = "";
            INVTA.TA044 = "";
            INVTA.TA045 = "";
            INVTA.TA046 = "";
            INVTA.TA047 = "";
            INVTA.TA048 = "";
            INVTA.TA049 = "0";
            INVTA.TA050 = "0";
            INVTA.TA051 = "";
            INVTA.TA052 = "";
            INVTA.TA053 = "";
            INVTA.TA054 = "";
            INVTA.TA055 = "0";
            INVTA.TA056 = "0";
            INVTA.TA057 = "0";
            INVTA.TA058 = "0";
            INVTA.TA059 = "";
            INVTA.TA060 = "";
            INVTA.TA061 = "";
            INVTA.TA062 = "";
            INVTA.TA063 = "";
            INVTA.TA064 = "";
            INVTA.TA065 = "";
            INVTA.TA066 = "";
            INVTA.TA067 = "";
            INVTA.TA068 = "";
            INVTA.TA200 = "";
            INVTA.UDF01 = INVTAUDF01;
            INVTA.UDF02 = "";
            INVTA.UDF03 = "";
            INVTA.UDF04 = "";
            INVTA.UDF05 = "";
            INVTA.UDF06 = "0";
            INVTA.UDF07 = "0";
            INVTA.UDF08 = "0";
            INVTA.UDF09 = "0";
            INVTA.UDF10 = "0";


            return INVTA;
        }


        public void CHECKBOMTD()
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
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT * FROM [TK].dbo.BOMTD WHERE UDF01='{0}'", textBoxID2.Text.ToString());
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
                    CHECKDELBOMMD = "Y";
                }
                else
                {
                    if (ds8.Tables["ds8"].Rows.Count >= 1)
                    {
                        CHECKDELBOMMD = "N";
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

        public void CHECKINVTA()
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
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT * FROM [TK].dbo.INVTA WHERE UDF01='{0}'", textBoxID3.Text.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter12 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder12 = new SqlCommandBuilder(adapter12);
                sqlConn.Open();
                ds12.Clear();
                adapter12.Fill(ds12, "ds12");
                sqlConn.Close();


                if (ds12.Tables["ds12"].Rows.Count == 0)
                {
                    CHECKDELINVTA = "Y";
                }
                else
                {
                    if (ds12.Tables["ds12"].Rows.Count >= 1)
                    {
                        CHECKDELINVTA = "N";
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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

        public void DELINVTARESLUT()
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKMK].[dbo].[INVTARESLUT]");
                sbSql.AppendFormat(" WHERE [SID]='{0}'", SID2);
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

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];

                    textBox13.Text = row.Cells["日期"].Value.ToString();
                    textBox14.Text = row.Cells["部門"].Value.ToString();
                    textBox15.Text = row.Cells["品號"].Value.ToString();
                    textBox16.Text = row.Cells["飲品"].Value.ToString();
                    textBox17.Text = row.Cells["其他"].Value.ToString();
                    textBox18.Text = row.Cells["數量"].Value.ToString();
                    textBox19.Text = row.Cells["原因"].Value.ToString();              
                    textBox20.Text = row.Cells["部門名"].Value.ToString();
                    textBoxID3.Text = row.Cells["ID"].Value.ToString();

                    TA003 = row.Cells["日期"].Value.ToString();
                    TA004 = row.Cells["部門"].Value.ToString();
                    TA005 = row.Cells["原因"].Value.ToString();
                    TA029= row.Cells["ID"].Value.ToString();
                    SID2 = row.Cells["ID"].Value.ToString();
                    INVTAUDF01 = row.Cells["ID"].Value.ToString();

                    CHECKINVTARESLUT();
                }
                else
                {                    
                    textBox13.Text = null;
                    textBox14.Text = null;
                    textBox15.Text = null;
                    textBox16.Text = null;
                    textBox17.Text = null;
                    textBox18.Text = null;
                    textBox19.Text = null;
                    textBox20.Text = null;
                    textBoxID3.Text = null;

                    TA003 = null;
                    TA004 = null;
                    TA005 = null;
                    TA029 = null;
                    SID2 = null;
                    INVTAUDF01 = null;
                }
            }
        }

        public void UPDATEINVTA()
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
               
                sbSql.AppendFormat(" UPDATE [TK].dbo.INVTA");
                sbSql.AppendFormat(" SET TA011=(SELECT SUM(TB007) FROM [TK].dbo.INVTB WHERE TB001='{0}' AND TB002='{1}') , TA012=(SELECT SUM(TB011) FROM [TK].dbo.INVTB WHERE TB001='{2}' AND TB002='{3}')",TA001,TA002,TA001,TA002);
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
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


        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];

                    textBox23.Text = row.Cells["代號"].Value.ToString();
                    textBox31.Text = row.Cells["名稱"].Value.ToString();

                    Search5();
                    Search6();
                }
                else
                {
                    textBox23.Text = null;
                    textBox31.Text = null;
                }
            }
        }

        public DataTable SEARCH_DRINKCREATOR()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            StringBuilder SQLQUERY1 = new StringBuilder();
            StringBuilder SQLQUERY2 = new StringBuilder();


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

                sbSql.AppendFormat(@" 
                                   SELECT
                                    [CREATOR]
                                    ,[CREATORNAMES]
                                    ,[USR_GROUP]
                                    ,[USR_GROUPNAMES]
                                    FROM [TKMK].[dbo].[DRINKCREATOR]
                                    ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    return ds.Tables["ds"];
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            if(textBox26.Text.Length>=6)
            {
                DataTable DT = DIND_VIEW_CardNo_MV004(textBox26.Text);

                if (DT != null && DT.Rows.Count >= 1)
                {                    
                    textBox27.Text = DT.Rows[0]["Name"].ToString();
                    textBox28.Text = DT.Rows[0]["CardNo"].ToString();
                    textBox24.Text = DT.Rows[0]["MV004"].ToString();
                    textBox33.Text = DT.Rows[0]["TITLE_NAME"].ToString();
                    comboBox3.SelectedValue = DT.Rows[0]["MV004"].ToString();
                }
                else
                {
                    //MessageBox.Show("此人員不存在");
                   
                    textBox27.Text = "";
                    textBox28.Text = "";
                    textBox33.Text = "";
                }
            }
          
        }
        private void textBox26_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox26.Text.Length >= 6)
            {
                DataTable DT = DIND_VIEW_CardNo_MV004(textBox26.Text);

                if (DT != null && DT.Rows.Count >= 1)
                {                   
                    textBox27.Text = DT.Rows[0]["Name"].ToString();
                    textBox28.Text = DT.Rows[0]["CardNo"].ToString();
                    textBox24.Text = DT.Rows[0]["MV004"].ToString();
                    textBox33.Text = DT.Rows[0]["TITLE_NAME"].ToString();
                    comboBox3.SelectedValue = DT.Rows[0]["MV004"].ToString();
                }
                else
                {
                    //MessageBox.Show("此人員不存在");
                  
                    textBox27.Text = "";
                    textBox28.Text = "";
                    textBox33.Text = "";
                }
            }
        }

        public DataTable DIND_VIEW_CardNo_MV004(string MV001)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            StringBuilder SQLQUERY1 = new StringBuilder();
            StringBuilder SQLQUERY2 = new StringBuilder();


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

                sbSql.AppendFormat(@" 
                                    SELECT
                                    [EmployeeID]
                                    ,[CardNo]
                                    ,[Name]
                                    ,[MV004]
                                    ,[TITLE_NAME]
                                    FROM [TKMK].[dbo].[VIEW_CardNo_MV004]
                                    WHERE ([EmployeeID] LIKE '%{0}%' OR [CardNo]  LIKE '%{0}%')
                                    ", MV001);

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    return ds.Tables["ds"];
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }
        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            textBox24.Text = null;
            if (!string.IsNullOrEmpty(comboBox3.SelectedValue.ToString()) && !comboBox3.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                textBox24.Text = comboBox3.SelectedValue.ToString();
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            label46.Text = comboBox4.SelectedValue.ToString();
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox27.Text))
            {
                textBox32.Text = comboBox5.SelectedValue.ToString() + "-" + textBox27.Text+ textBox33.Text;
            }
            
        }
        public void ADD_MKDRINKRECORD(
            string DATES
            , string MV001
            , string CARDNO
            , string NAMES
            , string DEP
            , string DEPNAME
            , string DRINKID
            , string DRINK
            , string OTHERS
            , string CUP
            , string REASON
            )
        {
            if (!string.IsNullOrEmpty(NAMES))
            {
                try
                {

                    //add ZWAREWHOUSEPURTH
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMK].[dbo].[MKDRINKRECORD]
                                        ([DATES],[MV001],[CARDNO],[NAMES],[DEP],[DEPNAME],[DRINKID],[DRINK],[OTHERS],[CUP],[REASON])                                       
                                        VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')
                                        ", DATES
                                        , MV001
                                        , CARDNO
                                        , NAMES
                                        , DEP
                                        , DEPNAME
                                        , DRINKID
                                        , DRINK
                                        , OTHERS
                                        , CUP
                                        , REASON);


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
        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            if (textBox34.Text.Length >= 6)
            {
                DataTable DT = DIND_VIEW_CardNo_MV004(textBox34.Text);

                if (DT != null && DT.Rows.Count >= 1)
                {
                    textBox35.Text = DT.Rows[0]["Name"].ToString();
                    textBox36.Text = DT.Rows[0]["CardNo"].ToString();
                    textBox1.Text = DT.Rows[0]["MV004"].ToString();                    
                    comboBox1.SelectedValue = DT.Rows[0]["MV004"].ToString();
                }
                else
                {
                    //MessageBox.Show("此人員不存在");
                    textBox35.Text = "";
                    textBox36.Text = "";
                    textBox1.Text = "";
                }
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
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

            UPDATE(
                        dateTimePicker3.Value.ToString("yyyMMdd")
                        , textBox34.Text
                        , textBox36.Text
                        , textBox35.Text
                        , textBox1.Text
                        , comboBox1.Text.ToString()
                        , comboBox2.Text.ToString()
                        , textBox2.Text
                        , textBox3.Text
                        , textBox4.Text
                        , DRINKID.Text
                        , textBoxID.Text
                        );

            SETSTAUSFIANL();

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            MessageBox.Show("完成");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //if (STATUS.Equals("EDIT"))
            //{
            //    UPDATE(
            //            dateTimePicker3.Value.ToString("yyyMMdd")
            //            , textBox34.Text
            //            , textBox36.Text
            //            , textBox35.Text
            //            , textBox1.Text
            //            , comboBox1.Text.ToString()
            //            , comboBox2.Text.ToString()
            //            , textBox2.Text
            //            , textBox3.Text
            //            , textBox4.Text
            //            , DRINKID.Text
            //            , textBoxID.Text
            //            );
            //}
            //else if (STATUS.Equals("ADD"))
            //{
            //    ADD();
            //}

            //STATUS = null;

            //SETSTAUSFIANL();

            //Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            //MessageBox.Show("完成");
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

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
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
            TD002 = GETMAXTD002(TD001,TD003);
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

                if(CHECKDELBOMMD.Equals("Y"))
                {
                    DELBOMTDRESLUT();
                    CHECKBOMTDRESLUT();
                }
                else if (CHECKDELBOMMD.Equals("N"))
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

        private void button11_Click(object sender, EventArgs e)
        {
            TA001 = "A111";
            TB012 = textBox21.Text.ToString();
            TA002 = GETMAXTA002(TA001,TA003);
            ADDINVTARESLUT(textBox14.Text,textBox13.Text,TA001,TA002);
            ADDINVTAB();
            UPDATEINVTA();

            CHECKINVTARESLUT();
        }
        private void button12_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHECKINVTA();

                if (CHECKDELINVTA.Equals("Y"))
                {
                    DELINVTARESLUT();
                    CHECKINVTARESLUT();
                }
                else if (CHECKDELINVTA.Equals("N"))
                {
                    MessageBox.Show("ERP還有費用單未刪除，請先刪除");
                }

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            Search4();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            TA001 = "A111";
            TA003 = dateTimePicker12.Value.ToString("yyyyMMdd");         
            TA002 = GETMAXTA002(TA001, TA003);
            TB012 = textBox22.Text;
            ADDINVTARESLUT(textBox23.Text, TA003, TA001, TA002);
            ADDINVTAB2();
            UPDATEINVTA();

            Search5();
            Search6();

            MessageBox.Show("已完成");
            //CHECKINVTARESLUT();
        }
        private void button15_Click(object sender, EventArgs e)
        {
            Search7(dateTimePicker13.Value.ToString("yyyyMMdd"));
        }

        private void button16_Click(object sender, EventArgs e)
        {
            ADD_MKDRINKRECORD(
             dateTimePicker14.Value.ToString("yyyy/MM/dd")
            , textBox26.Text
            , textBox28.Text
            , textBox27.Text
            , textBox24.Text
            , comboBox3.Text.ToString()
            , label46.Text
            , comboBox4.Text.ToString()
            , textBox29.Text
            , textBox30.Text
            , textBox32.Text
            );

             Search7(dateTimePicker13.Value.ToString("yyyyMMdd"));
        }







        #endregion

    
    }
}
