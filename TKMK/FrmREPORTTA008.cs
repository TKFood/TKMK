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
    /// </summary>
    public partial class FrmREPORTTA008 : Form
    {
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

        public void SETFASTREPORT(string SDATES, string EDATES, string TA008)
        {
            Report report1 = new Report();
            report1.Load(@"REPORT\\指定團客券報表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            //報表內有多個連線(Connection、Connection1、Connection2)，全部替換
            foreach (DataConnectionBase conn in report1.Dictionary.Connections)
            {
                conn.ConnectionString = sqlsb.ConnectionString;
            }

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SETSQL(SDATES, EDATES, TA008).ToString();

            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SETSQL1(SDATES, EDATES, TA008).ToString();

            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SETSQL2(SDATES, EDATES, TA008).ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        /// <summary>
        /// 交易筆數
        /// </summary>
        public StringBuilder SETSQL(string SDATES, string EDATES, string TA008)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            COUNT(*) AS '交易筆數'
                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            WHERE 1=1
                            AND TA001>='{0}' AND TA001<='{1}'
                            AND TA008 LIKE '{2}%'
                            ", SDATES, EDATES, TA008);

            return SB;
        }

        /// <summary>
        /// 總金額(含稅)
        /// </summary>
        public StringBuilder SETSQL1(string SDATES, string EDATES, string TA008)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            ISNULL(SUM(TA024),0) AS '總金額(含稅)'
                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            WHERE 1=1
                            AND TA001>='{0}' AND TA001<='{1}'
                            AND TA008 LIKE '{2}%'
                            ", SDATES, EDATES, TA008);

            return SB;
        }

        /// <summary>
        /// 商品明細
        /// </summary>
        public StringBuilder SETSQL2(string SDATES, string EDATES, string TA008)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            TB010 AS '品號'
                            ,MB002 AS '品名'
                            ,ISNULL(SUM(TB019),0) AS '銷售數量'
                            ,ISNULL(SUM((TB031+TB032)),0) AS '總金額(含稅)'
                            FROM [TK].dbo.POSTA WITH(NOLOCK),[TK].dbo.POSTB WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                            WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003 AND TA006=TB006
                            AND TB010=MB001
                            AND TA001>='{0}' AND TA001<='{1}'
                            AND TA008 LIKE '{2}%'
                            GROUP BY TB010,MB002
                            ORDER BY TB010,MB002
                            ", SDATES, EDATES, TA008);

            return SB;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            //避免單引號造成 SQL 錯誤/注入
            string TA008 = textBox1.Text.Trim().Replace("'", "''");

            if (string.IsNullOrEmpty(TA008))
            {
                MessageBox.Show("請輸入團客券號碼");
                return;
            }

            SETFASTREPORT(SDATES, EDATES, TA008);
        }
        #endregion
    }
}
