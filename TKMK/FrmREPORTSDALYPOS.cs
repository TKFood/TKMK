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
                                    ,[PUBNUMS] AS '試吃+公關數量'
                                    ,[NOWNUMS] AS '庫存數量'
                                    ,[COMMENTS] AS '備註'
                                    ,[ID]
                                    ,[CREATEDATES]
                                    FROM [TKMK].[dbo].[TBDAILYPOSTB]
                                    WHERE [SDATES]='{0}'
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["ID"].Value.ToString();
                    textBox2.Text = row.Cells["備註"].Value.ToString();
                }
            }

        }

        public void UPDATE_TBDAILYPOSTB_COMMENTS(string ID,string COMMENTS)
        {
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;
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
                                    UPDATE [TKMK].[dbo].[TBDAILYPOSTB]
                                    SET [COMMENTS]='{1}'
                                    WHERE [ID]='{0}'
                                    "
                                    , ID
                                    , COMMENTS
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

        public void ADD_TBDAILYPOSTB(string SDATES)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSql99 = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            DataTable DT_NOTIN = FIND_TBDAILYPOSTBNOTIN();
            if(DT_NOTIN!=null && DT_NOTIN.Rows.Count>=1)
            {
                foreach(DataRow DR in DT_NOTIN.Rows)
                {
                    sbSql99.AppendFormat(@" AND MB001 NOT LIKE '{0}%'", DR["MB001"].ToString().Trim());
                }
               
            }

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
                                    DELETE [TKMK].[dbo].[TBDAILYPOSTB]
                                    WHERE [SDATES]='{0}'

                                    --新增日期+品號
                                    INSERT INTO [TKMK].[dbo].[TBDAILYPOSTB]
                                    ([SDATES]
                                    ,[MB001]
                                    ,[MB002]
                                    )
                                    SELECT DISTINCT '{0}' , MB001, MB002
                                    FROM 
                                    (
                                        SELECT LA001 AS MB001, MB002 AS MB002
                                        FROM [TK].dbo.INVLA
                                        INNER JOIN [TK].dbo.INVMB ON LA001 = MB001
                                        WHERE (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                        AND LA009  IN ( '21002')
                                        AND LA004<='{0}'
                                        GROUP BY LA001, MB002
                                        HAVING SUM(LA005 * LA011) > 0

                                        UNION ALL

                                        SELECT TB010 AS MB001, MB002 AS MB002
                                        FROM [TK].dbo.POSTB
                                        INNER JOIN [TK].dbo.INVMB ON TB010 = MB001
                                        WHERE (TB010 LIKE '4%' OR TB010 LIKE '5%')
                                        AND TB002   IN ( '106702')
                                        AND TB001 = '{0}'
                                        GROUP BY TB010, MB002
                                        HAVING SUM(TB019) > 0

	                                    UNION ALL

	                                    SELECT LA001 AS MB001, MB002 AS MB002
                                        FROM [TK].dbo.INVLA
                                        INNER JOIN [TK].dbo.INVMB ON LA001 = MB001
                                        WHERE (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                        AND LA009  IN ( '21002')
	                                    AND LA005 IN (1)
                                        AND LA004='{0}'
                                        GROUP BY LA001, MB002
                                        HAVING SUM(LA005 * LA011) > 0

                                        UNION ALL

	                                    SELECT LA001 AS MB001, MB002 AS MB00
                                        FROM [TK].dbo.INVLA
                                        INNER JOIN [TK].dbo.INVMB ON LA001 = MB001
                                        WHERE (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                        AND LA009  IN ( '21002')
	                                    AND LA005 IN (-1)
	                                    AND LA006 IN ('A111')
                                        AND LA004='{0}'
                                        GROUP BY LA001, MB002
	                                    HAVING SUM(LA005 * LA011*-1)>0

                                    ) AS TEMP
                                    WHERE 1=1
                                    {1}
                                    ORDER BY MB001, MB002

                                    --更新庫存量
                                    UPDATE [TKMK].[dbo].[TBDAILYPOSTB]
                                    SET [NOWNUMS]=TEMP.NUMS
                                    FROM 
                                    (
	                                    SELECT LA001,MB002,SUM(LA005*LA011) AS NUMS
	                                    FROM [TK].dbo.INVLA,[TK].dbo.INVMB
	                                    WHERE LA001=MB001
	                                    AND (LA001 LIKE '4%' OR LA005 LIKE '5%')
	                                    AND LA009 IN ('21002')
                                        AND LA004<='{0}'
	                                    GROUP BY  LA001,MB002
                                    HAVING SUM(LA005*LA011)>0
                                    ) AS TEMP
                                    WHERE TEMP.LA001=[TBDAILYPOSTB].MB001
                                    AND [TBDAILYPOSTB].[SDATES]='{0}'

                                    --更新銷售量
                                    UPDATE [TKMK].[dbo].[TBDAILYPOSTB]
                                    SET [SALENUMS]=TEMP.TB019
                                    FROM 
                                    (
	                                    SELECT TB010,MB002,SUM(TB019) TB019
	                                    FROM [TK].dbo.POSTB,[TK].dbo.INVMB
	                                    WHERE TB010=MB001
	                                    AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
	                                    AND  TB002 IN ('106702')
	                                    AND TB001='{0}'
	                                    GROUP BY TB010,MB002
	                                    HAVING SUM(TB019)>0
                                    ) AS TEMP
                                    WHERE TEMP.TB010=[TBDAILYPOSTB].MB001
                                    AND [TBDAILYPOSTB].[SDATES]='{0}'

                                    --更新進貨量
                                    UPDATE [TKMK].[dbo].[TBDAILYPOSTB]
                                    SET INNUMS=TEMP.NUMS
                                    FROM 
                                    (
	                                    SELECT LA001, MB002,SUM(LA005*LA011) AS NUMS
                                        FROM [TK].dbo.INVLA
                                        INNER JOIN [TK].dbo.INVMB ON LA001 = MB001
                                        WHERE (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                        AND LA009  IN ( '21002')
	                                    AND LA005 IN (1)
                                        AND LA004='{0}'
                                        GROUP BY LA001, MB002
                                        HAVING SUM(LA005 * LA011) > 0
                                    ) AS TEMP
                                    WHERE TEMP.LA001=[TBDAILYPOSTB].MB001
                                    AND [TBDAILYPOSTB].[SDATES]='{0}'

                                    --更新 試吃+公關
                                    UPDATE [TKMK].[dbo].[TBDAILYPOSTB]
                                    SET PUBNUMS=TEMP.NUMS
                                    FROM 
                                    (
	                                    SELECT LA001 AS MB001, MB002 AS MB002,SUM(LA005 * LA011*-1)  AS NUMS
                                        FROM [TK].dbo.INVLA
                                        INNER JOIN [TK].dbo.INVMB ON LA001 = MB001
                                        WHERE (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                        AND LA009  IN ( '21002')
	                                    AND LA005 IN (-1)
	                                    AND LA006 IN ('A111')
                                        AND LA004='{0}'
                                        GROUP BY LA001, MB002
	                                    HAVING SUM(LA005 * LA011*-1)>0
 
                                    ) AS TEMP
                                    WHERE TEMP.MB001=[TBDAILYPOSTB].MB001
                                    AND [TBDAILYPOSTB].[SDATES]='{0}'
                                    "
                                    , SDATES, sbSql99.ToString()
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

        public DataTable FIND_TBDAILYPOSTBNOTIN()
        {
            DataTable DT = new DataTable();

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
                                    [MB001]
                                    FROM [TKMK].[dbo].[TBDAILYPOSTBNOTIN]
                                                                        
                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
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

        public void SETFASTREPORT(string SDATES)
        {
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();
            report1.Load(@"REPORT\硯微墨每日商品統計表.frx");

            SQL = SETSQL(SDATES);
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();
            report1.SetParameterValue("P1", SDATES);


            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATES)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"    
                            SELECT 
                            [SDATES] AS '日期'
                            ,[MB001] AS '品號'
                            ,[MB002] AS '品名'
                            ,[SALENUMS] AS '銷售數量'
                            ,[INNUMS] AS '入庫數量'
                            ,[PUBNUMS] AS '試吃+公關數量'
                            ,[NOWNUMS] AS '庫存數量'
                            ,[COMMENTS] AS '備註'
                            ,[ID]
                            ,[CREATEDATES]
                            FROM [TKMK].[dbo].[TBDAILYPOSTB]
                            WHERE [SDATES]='{0}'
                            ORDER BY [MB001]
                            ", SDATES);
            SB.AppendFormat(@" ");

            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");

            MESSAGESHOW MSGSHOW = new MESSAGESHOW();
            // 鎖定控制項
            this.Enabled = false;
            // 顯示跳出視窗
            MSGSHOW.Show();

            // 使用非同步操作執行長時間運行的操作
            Task.Run(() =>
            {
                // ADD_TBDAILYPOSTB
                ADD_TBDAILYPOSTB(SDATES);

                // 更新 UI，確保在主 UI 線程上執行
                Invoke(new Action(() =>
                {
                    SEARCHGROUPSALES(SDATES);
                    MSGSHOW.Close();
                    // 解除鎖定
                    this.Enabled = true;
                }));
            });
            
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text) )
            {
                UPDATE_TBDAILYPOSTB_COMMENTS(textBox1.Text, textBox2.Text.Trim());

                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
                textBox2.Text = null;
                MessageBox.Show("完成", "完成");
            }
           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            SETFASTREPORT(SDATES);
        }


        #endregion


    }
}
