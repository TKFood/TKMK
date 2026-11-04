using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TKITDLL;

namespace TKMK
{
    /// <summary>
    /// 團務佣金計算 (20261104 新增)
    /// 參考 frmGROUPSALESBYTA008 的 button18 / SETMONEYS_NEW，計算規則完全相同，差別在做法：
    /// 1. 只算「預約接團」的車 (原本會全部算完，但只有預約接團會被 UPDATE，其他白算)
    /// 2. 程式先查出當天要用的「日期、業務員帳號」，用實際值當參數 + OPTION(RECOMPILE) 抓 POS
    ///    (原本用變數/子查詢，SQL Server 猜不到筆數，容易選到掃整張 POSTA/POSTB 的執行計畫)
    /// 3. POS 資料抓進 #TA/#TB/#TC 後，原本的佣金公式只對這些小暫存表計算
    /// 4. 交易(Transaction) 只包住最後兩個 UPDATE，鎖 GROUPSALES 的時間最短
    /// 5. 每一步都計時，顯示在右邊的紀錄裡，可以看出到底慢在哪一步
    /// 6. 全部在背景執行，畫面不會卡住；出錯會顯示訊息
    /// </summary>
    public partial class frmGROUPSALESCOMMISSION : Form
    {
        public frmGROUPSALESCOMMISSION()
        {
            InitializeComponent();
        }

        private void frmGROUPSALESCOMMISSION_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
        }

        #region FUNCTION

        /// <summary>
        /// 取得解密後的連線字串
        /// </summary>
        private string GETCONNSTRING()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            return sqlsb.ConnectionString;
        }

        /// <summary>
        /// 計算佣金 (背景執行)：呼叫共用的 GROUPSALESCOMMISSIONCAL
        /// </summary>
        public List<string> CALMONEYS(string SDATES)
        {
            return GROUPSALESCOMMISSIONCAL.CALMONEYS(GETCONNSTRING(), SDATES);
        }


        /// <summary>
        /// 查當天來車資料 (顯示用)
        /// </summary>
        public DataTable SEARCHGROUPSALES(string SDATES)
        {
            DateTime D1 = DateTime.ParseExact(SDATES, "yyyyMMdd", null);

            List<SqlParameter> P = new List<SqlParameter>();
            P.Add(new SqlParameter("@D1", SqlDbType.DateTime) { Value = D1 });
            P.Add(new SqlParameter("@D2", SqlDbType.DateTime) { Value = D1.AddDays(1) });

            using (SqlConnection conn = new SqlConnection(GETCONNSTRING()))
            {
                conn.Open();
                return GROUPSALESCOMMISSIONCAL.QUERY(conn, null, @"
                                    SELECT  
                                    [SERNO] AS '序號'
                                    ,[CARNAME] AS '車名'
                                    ,[CARNO] AS '車號'
                                    ,[CARKIND] AS '車種'
                                    ,[ISEXCHANGE] AS '兌換券'
                                    ,[STATUS] AS '狀態'
                                    ,[CARNUM] AS '車數'
                                    ,[GUSETNUM] AS '交易筆數'
                                    ,[SALESMMONEYS] AS '消費總額'
                                    ,[EXCHANGEMONEYS] AS '領券額'
                                    ,[EXCHANGETOTALMONEYS] AS '券總額'
                                    ,[EXCHANGESALESMMONEYS] AS '券消費'
                                    ,[SPECIALMNUMS] AS '特賣數'
                                    ,[SPECIALMONEYS] AS '特賣獎金'
                                    ,[COMMISSIONBASEMONEYS] AS '茶水費'
                                    ,[COMMISSIONPCT] AS '抽佣比率'
                                    ,[COMMISSIONPCTMONEYS] AS '消費獎金'
                                    ,[TOTALCOMMISSIONMONEYS] AS '總獎金'
                                    ,[TA008NO] AS '業務員名'
                                    ,[TA008] AS '業務員帳號'
                                    ,CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'
                                    ,[ID]
                                    FROM [TKMK].[dbo].[GROUPSALES] WITH(NOLOCK)
                                    WHERE [CREATEDATES]>=@D1 AND [CREATEDATES]<@D2
                                    AND [STATUS]<>'取消預約'
                                    ORDER BY CONVERT(int,[SERNO]) DESC
                                    ", P);
            }
        }

        private void SHOWGRID(DataTable dt)
        {
            dataGridView1.DataSource = dt;

            foreach (string col in new[] { "消費總額", "領券額", "券總額", "券消費", "特賣數", "特賣獎金", "茶水費", "消費獎金", "總獎金", "交易筆數", "車數" })
            {
                if (dataGridView1.Columns.Contains(col))
                {
                    dataGridView1.Columns[col].DefaultCellStyle.Format = "#,##0";
                    dataGridView1.Columns[col].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }

            dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
        }

        private void LOG(string msg)
        {
            textBox1.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + msg + Environment.NewLine);
        }

        #endregion

        #region BUTTON

        /// <summary>
        /// 計算佣金
        /// </summary>
        private async void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");

            button1.Enabled = false;
            button2.Enabled = false;
            Cursor = Cursors.WaitCursor;
            LOG("===== 開始計算 " + SDATES + " =====");

            try
            {
                List<string> LOGS = await Task.Run(() => CALMONEYS(SDATES));
                foreach (string s in LOGS)
                {
                    LOG(s);
                }

                Stopwatch SW = Stopwatch.StartNew();
                DataTable dt = await Task.Run(() => SEARCHGROUPSALES(SDATES));
                SHOWGRID(dt);
                LOG(string.Format("7.查詢來車資料：{0} 筆，{1:0.00} 秒", dt.Rows.Count, SW.Elapsed.TotalSeconds));

                label2.Text = "更新時間 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            }
            catch (Exception ex)
            {
                LOG("失敗：" + ex.Message);
                MessageBox.Show("計算佣金失敗：" + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                button1.Enabled = true;
                button2.Enabled = true;
            }
        }

        /// <summary>
        /// 只查詢，不計算
        /// </summary>
        private async void button2_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");

            button1.Enabled = false;
            button2.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                Stopwatch SW = Stopwatch.StartNew();
                DataTable dt = await Task.Run(() => SEARCHGROUPSALES(SDATES));
                SHOWGRID(dt);
                LOG(string.Format("查詢來車資料 {0}：{1} 筆，{2:0.00} 秒", SDATES, dt.Rows.Count, SW.Elapsed.TotalSeconds));
            }
            catch (Exception ex)
            {
                LOG("失敗：" + ex.Message);
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                button1.Enabled = true;
                button2.Enabled = true;
            }
        }

        #endregion
    }
}
