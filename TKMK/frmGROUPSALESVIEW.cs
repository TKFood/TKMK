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
    public partial class frmGROUPSALESVIEW : Form
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
        int ROWSINDEX = 0;
        int COLUMNSINDEX = 0;

        public frmGROUPSALESVIEW()
        {
            InitializeComponent();


            dateTimePicker1.Value = DateTime.Now;


            timer1.Enabled = true;
            timer1.Interval = 1000 * 30;
            timer1.Start();
        }

        #region FUNCTION



        private void timer1_Tick(object sender, EventArgs e)
        {

            if (STATUSCONTROLLER.Equals("VIEW"))
            {
                dateTimePicker1.Value = GETDBDATES();
                
                label29.Text = "";
                label29.Text = "更新時間" + dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");


                SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));
             
            }
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
                sbSql.AppendFormat(@"  [SERNO] AS '序號',[CARNAME] AS '車名',[CARNO] AS '車號',[CARKIND] AS '車種',[GROUPKIND]  AS '團類',[ISEXCHANGE] AS '兌換券',[EXCHANGETOTALMONEYS] AS '券總額',[EXCHANGESALESMMONEYS] AS '券消費',[SALESMMONEYS] AS '消費總額'");
                sbSql.AppendFormat(@"  ,[SPECIALMNUMS] AS '特賣數',[SPECIALMONEYS] AS '特賣獎金',[COMMISSIONBASEMONEYS] AS '茶水費',[COMMISSIONPCTMONEYS] AS '消費獎金',[TOTALCOMMISSIONMONEYS] AS '總獎金',[CARNUM] AS '車數',[GUSETNUM] AS '來客數',[EXCHANNO] AS '優惠券名',[EXCHANACOOUNT] AS '優惠券帳號',CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間',CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間',[STATUS] AS '狀態'");
                sbSql.AppendFormat(@"  ,CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間',CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間',[COMMISSIONPCT] AS '抽佣比率',[EXCHANGEMONEYS] AS '領券額',[ID],[CREATEDATES]");
                sbSql.AppendFormat(@"  FROM [TKMK].[dbo].[GROUPSALES]");
                sbSql.AppendFormat(@"  WHERE CONVERT(nvarchar,[CREATEDATES],112)='{0}' ", CREATEDATES);
                sbSql.AppendFormat(@"  AND [STATUS] NOT IN ('取消預約') ");
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
                        dataGridView1.Columns["序號"].Width = 30;
                        dataGridView1.Columns["車名"].Width = 80;
                        dataGridView1.Columns["車號"].Width = 100;
                        dataGridView1.Columns["車種"].Width = 40;
                        dataGridView1.Columns["團類"].Width = 80;
                        dataGridView1.Columns["兌換券"].Width = 20;

                        dataGridView1.Columns["券總額"].Width = 60;
                        dataGridView1.Columns["券消費"].Width = 60;
                        dataGridView1.Columns["消費總額"].Width = 60;
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
                            dgRow.Cells["券總額"].Style.Font = new Font("Tahoma", 14);
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


                if (ROWSINDEX > 0 || COLUMNSINDEX > 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[ROWSINDEX].Cells[COLUMNSINDEX];

                    DataGridViewRow row = dataGridView1.Rows[ROWSINDEX];
                    

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
            SB.AppendFormat(@" WHERE CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112)>='{0}' AND CONVERT(NVARCHAR,[PURGROUPSTARTDATES],112)<='{1}'", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@" AND [STATUS]='完成接團'");
            SB.AppendFormat(@" ORDER BY CONVERT(NVARCHAR,[PURGROUPSTARTDATES], 112),[SERNO]");
            SB.AppendFormat(@"  ");

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

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHGROUPSALES(dateTimePicker1.Value.ToString("yyyyMMdd"));           

            label29.Text = "";
            label29.Text = "更新時間" + dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm:ss");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion


    }
}
