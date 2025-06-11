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
    public partial class FrmREPORTSCOUNTS : Form
    {
        public FrmREPORTSCOUNTS()
        {
            InitializeComponent();
        }

        #region FUNCTION       
        private void FrmREPORTSCOUNTS_Load(object sender, EventArgs e)
        {
            SETDATES();
        }

        public void SETDATES()
        {
            // 本月第一天
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            // 本月最後一天
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            dateTimePicker1.Value = firstDayOfMonth;
            dateTimePicker2.Value = lastDayOfMonth;
        }

        public void SETFASTREPORT(string DATES_START, string DATES_END)
        {
            SqlConnection sqlConn = new SqlConnection();
             
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            Report report1 = new Report();
            report1.Load(@"REPORT\消費筆數.frx");

            SQL1= SETSQL1(DATES_START, DATES_END);
            SQL2 = SETSQL2(DATES_START, DATES_END);
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
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table1.SelectCommand = SQL2.ToString();
            //report1.SetParameterValue("P1", SDATES);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string DATES_START,string DATES_END)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                               '{0}' AS '查詢日期',
                                交易店號,
	                            MA002 AS '門市',
                                SUM(總交易筆數) 總交易筆數,
                                SUM(金額500元的交易筆數) 金額500元的交易筆數,
                                SUM(金額1000元的交易筆數) 金額1000元的交易筆數,
                                CAST(SUM(金額500元的交易筆數) AS DECIMAL(10,2)) / NULLIF(SUM(總交易筆數), 0) AS [金額500元的%],
                                CAST(SUM(金額1000元的交易筆數) AS DECIMAL(10,2)) / NULLIF(SUM(總交易筆數), 0) AS [金額1000元的%]
                            FROM 
                            (
                                SELECT 
                                    TA001 AS 交易日期,
                                    TA002 AS 交易店號,
                                    COUNT(*) AS 總交易筆數,
                                    (
                                        SELECT COUNT(*) 
                                        FROM [TK].dbo.POSTA TA1 
                                        WHERE TA1.TA001 = POSTA.TA001 
                                          AND TA1.TA002 = POSTA.TA002  
                                          AND TA1.TA026 >= 500
                                    ) AS 金額500元的交易筆數,
                                    (
                                        SELECT COUNT(*) 
                                        FROM [TK].dbo.POSTA TA1 
                                        WHERE TA1.TA001 = POSTA.TA001 
                                          AND TA1.TA002 = POSTA.TA002  
                                          AND TA1.TA026 >= 1000
                                    ) AS 金額1000元的交易筆數
                                FROM [TK].dbo.POSTA WITH(NOLOCK)
                                WHERE TA002 IN (
                                    SELECT TA002
                                    FROM [TKMK].[dbo].[TB_POS_TA002]
                                )
                                AND TA001 >= '{1}' AND TA001 <= '{2}'
                                GROUP BY TA001, TA002
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=交易店號
                            GROUP BY 交易店號,MA002
                            ORDER BY 交易店號;
                            ", DATES_START+"~"+DATES_END, DATES_START, DATES_END);
            SB.AppendFormat(@" ");

            return SB;

        }
        public StringBuilder SETSQL2(string DATES_START, string DATES_END)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                                交易日期,
                                交易店號,
	                            MA002 AS '門市',
                                總交易筆數,
                                金額500元的交易筆數,
                                金額1000元的交易筆數,
                                CAST(金額500元的交易筆數 AS DECIMAL(10,2)) / NULLIF(總交易筆數, 0) AS [金額500元的%],
                                CAST(金額1000元的交易筆數 AS DECIMAL(10,2)) / NULLIF(總交易筆數, 0) AS [金額1000元的%]
                            FROM 
                            (
                                SELECT 
                                    TA001 AS 交易日期,
                                    TA002 AS 交易店號,
                                    COUNT(*) AS 總交易筆數,
                                    (
                                        SELECT COUNT(*) 
                                        FROM [TK].dbo.POSTA TA1 
                                        WHERE TA1.TA001 = POSTA.TA001 
                                          AND TA1.TA002 = POSTA.TA002  
                                          AND TA1.TA026 >= 500
                                    ) AS 金額500元的交易筆數,
                                    (
                                        SELECT COUNT(*) 
                                        FROM [TK].dbo.POSTA TA1 
                                        WHERE TA1.TA001 = POSTA.TA001 
                                          AND TA1.TA002 = POSTA.TA002  
                                          AND TA1.TA026 >= 1000
                                    ) AS 金額1000元的交易筆數
                                FROM [TK].dbo.POSTA WITH(NOLOCK)
                                WHERE TA002 IN (
                                    SELECT TA002
                                    FROM [TKMK].[dbo].[TB_POS_TA002]
                                )
                                AND TA001 >= '{0}' AND TA001 <= '{1}'
                                GROUP BY TA001, TA002
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=交易店號
                            ORDER BY 交易店號,交易日期
                            ", DATES_START, DATES_END);
            SB.AppendFormat(@" ");

            return SB;

        }

        #endregion

        #region BUTTON

        private void button3_Click(object sender, EventArgs e)
        {
            string DATES_START = dateTimePicker1.Value.ToString("yyyyMMdd");
            string DATES_END = dateTimePicker2.Value.ToString("yyyyMMdd");
            SETFASTREPORT(DATES_START, DATES_END);
        }

        #endregion

      
    }
}
