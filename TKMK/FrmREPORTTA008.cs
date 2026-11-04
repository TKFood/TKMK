using FastReport;
using FastReport.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TKITDLL;

namespace TKMK
{
    /// <summary>
    /// 指定團客券報表
    /// 依 POSTA.TA008 (團客券號碼) 開頭 + 交易日期區間 統計
    ///
    /// 20261104 效能改善：
    /// 1. 原本報表 3 個 TableDataSource 各自查 DB，POSTA 被掃 3 次；
    ///    改為程式端用「一個批次」先把符合條件的 POSTA 篩進 #TA，再由 #TA 算出 3 個結果 (POSTA 只掃 1 次)
    /// 2. 商品明細 JOIN POSTB 時，明確加上 TB001 日期區間，讓 POSTB 可用索引 seek，不必整表掃
    /// 3. 先依品號彙總，再 JOIN INVMB 取品名 (JOIN 的筆數大幅減少)
    /// 4. 查詢改在背景執行，畫面不會「沒有回應」；CommandTimeout 放寬為 300 秒
    /// 5. 參數化查詢
    /// 查完的 DataTable 用 RegisterData 交給報表，frx 版面/欄位名稱都不用改
    /// </summary>
    public partial class FrmREPORTTA008 : Form
    {
        //SQL 逾時秒數
        private const int SQL_TIMEOUT = 300;

        public FrmREPORTTA008()
        {
            InitializeComponent();
        }

        private void FrmREPORTTA008_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            dateTimePicker2.Value = DateTime.Now;
            textBox1.Text = "68";
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
        /// LIKE 用：跳脫 % _ [ 這些萬用字元，只做「開頭符合」
        /// </summary>
        private static string ESCAPELIKE(string s)
        {
            return s.Replace("[", "[[]").Replace("%", "[%]").Replace("_", "[_]");
        }

        /// <summary>
        /// 一次查詢取得 3 個結果
        /// Tables[0] = 交易筆數 (Table)
        /// Tables[1] = 總金額(含稅) (Table1)
        /// Tables[2] = 商品明細 (Table2)
        /// </summary>
        public DataSet SEARCHDATA(string connString, string SDATES, string EDATES, string TA008)
        {
            DataSet ds = new DataSet();

            using (SqlConnection conn = new SqlConnection(connString))
            using (SqlCommand cmd = new SqlCommand(SETSQL().ToString(), conn))
            {
                cmd.CommandTimeout = SQL_TIMEOUT;
                cmd.Parameters.Add("@SDATE", SqlDbType.NVarChar, 20).Value = SDATES;
                cmd.Parameters.Add("@EDATE", SqlDbType.NVarChar, 20).Value = EDATES;
                cmd.Parameters.Add("@TA008", SqlDbType.NVarChar, 100).Value = ESCAPELIKE(TA008) + "%";

                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(ds);
                }
            }

            return ds;
        }

        /// <summary>
        /// 一個批次：POSTA 只掃一次放進 #TA，再算 交易筆數 / 總金額 / 商品明細
        /// </summary>
        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.Append(@"
                        SET NOCOUNT ON;

                        IF OBJECT_ID('tempdb..#TA') IS NOT NULL DROP TABLE #TA;

                        --符合條件的交易單頭，只掃 POSTA 一次
                        SELECT TA001,TA002,TA003,TA006,TA024
                        INTO #TA
                        FROM [TK].dbo.POSTA WITH(NOLOCK)
                        WHERE TA001>=@SDATE AND TA001<=@EDATE
                        AND TA008 LIKE @TA008;

                        CREATE CLUSTERED INDEX IX_TA ON #TA (TA001,TA002,TA003,TA006);

                        --Table：交易筆數
                        SELECT COUNT(*) AS [交易筆數]
                        FROM #TA;

                        --Table1：總金額(含稅)
                        SELECT ISNULL(SUM(TA024),0) AS [總金額(含稅)]
                        FROM #TA;

                        --Table2：商品明細
                        --TB001 也加日期區間，POSTB 才能用索引範圍搜尋；先彙總再 JOIN INVMB
                        ;WITH TB AS
                        (
                            SELECT POSTB.TB010
                            ,SUM(POSTB.TB019) AS QTY
                            ,SUM(POSTB.TB031+POSTB.TB032) AS AMT
                            FROM #TA
                            INNER JOIN [TK].dbo.POSTB WITH(NOLOCK)
                                ON POSTB.TB001=#TA.TA001 AND POSTB.TB002=#TA.TA002
                                AND POSTB.TB003=#TA.TA003 AND POSTB.TB006=#TA.TA006
                            WHERE POSTB.TB001>=@SDATE AND POSTB.TB001<=@EDATE
                            GROUP BY POSTB.TB010
                        )
                        SELECT
                        TB.TB010 AS [品號]
                        ,INVMB.MB002 AS [品名]
                        ,ISNULL(TB.QTY,0) AS [銷售數量]
                        ,ISNULL(TB.AMT,0) AS [總金額(含稅)]
                        FROM TB
                        INNER JOIN [TK].dbo.INVMB WITH(NOLOCK) ON INVMB.MB001=TB.TB010
                        ORDER BY TB.TB010,INVMB.MB002;

                        DROP TABLE #TA;
                        ");

            return SB;
        }

        /// <summary>
        /// 把查好的資料交給報表 (不再讓報表自己連 DB 重查)
        /// </summary>
        public void SETFASTREPORT(DataSet ds)
        {
            Report report1 = new Report();
            report1.Load(@"REPORT\\指定團客券報表.frx");

            //記下每個 DataBand 原本綁的資料來源名稱 (Table / Table1 / Table2)
            Dictionary<DataBand, string> BANDS = new Dictionary<DataBand, string>();
            foreach (Base obj in report1.AllObjects)
            {
                DataBand band = obj as DataBand;
                if (band != null && band.DataSource != null)
                {
                    BANDS[band] = band.DataSource.Name;
                }
            }

            //移除報表內的 SQL 連線 (Connection、Connection1、Connection2)，改用程式查好的資料
            report1.Dictionary.Connections.Clear();

            string[] NAMES = { "Table", "Table1", "Table2" };
            for (int i = 0; i < NAMES.Length && i < ds.Tables.Count; i++)
            {
                ds.Tables[i].TableName = NAMES[i];
                report1.RegisterData(ds.Tables[i], NAMES[i]);

                DataSourceBase source = report1.GetDataSource(NAMES[i]);
                if (source != null)
                {
                    source.Enabled = true;
                }
            }

            //DataBand 重新綁定到新的資料來源
            foreach (KeyValuePair<DataBand, string> kv in BANDS)
            {
                DataSourceBase source = report1.GetDataSource(kv.Value);
                if (source != null)
                {
                    kv.Key.DataSource = source;
                }
            }

            report1.Preview = previewControl1;
            report1.Show();
        }

        #endregion

        #region BUTTON
        private async void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            string TA008 = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(TA008))
            {
                MessageBox.Show("請輸入團客券號碼");
                return;
            }

            button1.Enabled = false;
            string oldText = button1.Text;
            button1.Text = "查詢中...";
            Cursor = Cursors.WaitCursor;

            try
            {
                string connString = GETCONNSTRING();

                //SQL 在背景執行，畫面不會凍結
                DataSet ds = await Task.Run(() => SEARCHDATA(connString, SDATES, EDATES, TA008));

                SETFASTREPORT(ds);
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                button1.Text = oldText;
                button1.Enabled = true;
            }
        }
        #endregion
    }
}
