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


        public frmGROUPSALESMODIFY()
        {
            InitializeComponent();
        }


        #region FUNCTION       
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {   
            //查詢本日來車資料
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));

        }
        #endregion


    }
}
