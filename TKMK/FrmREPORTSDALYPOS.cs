using System;
using System.Collections.Generic;
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
    public partial class FrmREPORTSDALYPOS : Form
    {
        SqlConnection sqlConn = new SqlConnection();

        public FrmREPORTSDALYPOS()
        {
            InitializeComponent();
        }

        #region FUNCTION       
        private void FrmREPORTSDALYPOS_Load(object sender, EventArgs e)
        {
            SETDATES();
        }

        public void SETDATES()
        {
            dateTimePicker1.Value = DateTime.Now;
        }
        public void SEARCHGROUPSALES(string SDATES)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

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
                                    [SDATES] AS '日期'
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[SALENUMS] AS '銷售數量'
                                    ,[INNUMS] AS '入庫數量'
                                    ,[NOWNUMS] AS '庫存數量'
                                    ,[COMMENTS] AS '備註'
                                    ,[ID]
                                    ,[CREATEDATES]
                                    FROM [TKMK].[dbo].[TBDAILYPOSTB]
                                    WHERE [SDATES]='20250326'
                                    ORDER BY [MB001]
                                                                        
                                    ", SDATES);


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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }


        #endregion

       
    }
}
