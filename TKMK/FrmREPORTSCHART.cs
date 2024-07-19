using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Globalization;

namespace TKMK
{
    public partial class FrmREPORTSCHART : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        SqlTransaction tran;
        int result;

        public FrmREPORTSCHART()
        {
            InitializeComponent();

            SETDATES();
        }

        public void SETDATES()
        {
            DateTime FirstDay = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);

            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = LastDay;
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\團車業績圖表.frx");

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
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                              
                            SELECT *
                            FROM (
                            SELECT
                            '全部' AS '分類'
                            ,YEAR([CREATEDATES]) AS '年'
                            ,MONTH([CREATEDATES]) AS '月份'
                            ,COUNT([CARNUM])  AS  '車數'
                            ,SUM([SALESMMONEYS])  AS  '總團車銷售金額'
                            ,SUM([COMMISSIONBASEMONEYS])  AS  '總茶水費'
                            ,SUM([COMMISSIONPCTMONEYS])      AS  '總佣金' 
                            ,SUM([TOTALCOMMISSIONMONEYS])  AS  '總佣金+總茶水費'
                            ,(
                            SELECT SUM(TA026) 
                            FROM [TK].dbo.POSTA
                            WHERE TA002 LIKE '1067%'
                            AND YEAR(TA001)=YEAR([CREATEDATES]) AND MONTH(TA001)=MONTH([CREATEDATES])
                            ) AS  '觀光+硯微墨的總銷售金額'

                            FROM [TKMK].[dbo].[GROUPSALES]
                            WHERE [STATUS]='完成接團'
                            AND CONVERT(NVARCHAR,[CREATEDATES],112)>='{0}' AND CONVERT(NVARCHAR,[CREATEDATES],112)<='{1}' 
                            GROUP BY YEAR([CREATEDATES]),MONTH([CREATEDATES])
                            UNION ALL

                            SELECT
                            '滿5000元以上' AS '分類'
                            ,YEAR([CREATEDATES]) AS '年'
                            ,MONTH([CREATEDATES]) AS '月份'
                            ,COUNT([CARNUM])  AS  '車數'
                            ,SUM([SALESMMONEYS])  AS  '總團車銷售金額'
                            ,SUM([COMMISSIONBASEMONEYS])  AS  '總茶水費'
                            ,SUM([COMMISSIONPCTMONEYS])      AS  '總佣金' 
                            ,SUM([TOTALCOMMISSIONMONEYS])  AS  '總佣金+總茶水費'
                           ,(
                            SELECT SUM(TA026) 
                            FROM [TK].dbo.POSTA
                            WHERE TA002 LIKE '1067%'
                            AND YEAR(TA001)=YEAR([CREATEDATES]) AND MONTH(TA001)=MONTH([CREATEDATES])
                            ) AS  '觀光+硯微墨的總銷售金額'

                            FROM [TKMK].[dbo].[GROUPSALES]
                            WHERE [STATUS]='完成接團'
                            AND CONVERT(NVARCHAR,[CREATEDATES],112)>='{0}' AND CONVERT(NVARCHAR,[CREATEDATES],112)<='{1}' 
                            AND [SALESMMONEYS]>=5000
                            GROUP BY YEAR([CREATEDATES]),MONTH([CREATEDATES])
                            ) AS TEMP
                            ORDER BY 分類,年,月份

   

                            ", SDATES, EDATES);

            return SB;

        }


        #region FUNCTION

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        #endregion
    }
}
