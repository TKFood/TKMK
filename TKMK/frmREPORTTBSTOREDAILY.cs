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
namespace TKMK
{
    public partial class frmREPORTTBSTOREDAILY : Form
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

        public frmREPORTTBSTOREDAILY()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDATE, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATE, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\0901.門市營業日誌.frx");

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

        public StringBuilder SETSQL(string SDATE,string EDATES)
        {
            StringBuilder SB = new StringBuilder();

         
            SB.AppendFormat(@" 

                            SELECT 
                            [ID] 
                            ,[FIELD1] AS '門店'
                            ,[FIELD2] AS '上班人員'
                            ,[FIELD3] AS '早班(A班)'
                            ,[FIELD4] AS '午班(B班)'
                            ,[FIELD5] AS '日期'
                            ,[FIELD6] AS '星期'
                            ,[FIELD7] AS '天氣'
                            ,[FIELD8] AS '早午班－交班紀錄'
                            ,[FIELD9] AS '早班-服儀確認'
                            ,[FIELD10] AS '早班-設備開啟'
                            ,[FIELD11] AS '早班-清潔檢查'
                            ,[FIELD12] AS '早班-商品陳列檢查'
                            ,[FIELD13] AS '早班-廣告檢查'
                            ,[FIELD14] AS '早班-發票號碼'
                            ,[FIELD15] AS '早班-額外金額'
                            ,[FIELD16] AS '早班-清點零用金'
                            ,[FIELD17] AS '早班-短溢原因'
                            ,[FIELD18] AS '早班-零用金短／溢'
                            ,[FIELD19] AS 'FIELD19'
                            ,[FIELD20] AS '早班-交接事項備註'
                            ,[FIELD21] AS '早班-事件記錄'
                            ,[FIELD22] AS 'FIELD22'
                            ,[FIELD23] AS '午班－閉店紀錄'
                            ,[FIELD24] AS '午班-服儀確認'
                            ,[FIELD25] AS '午班-商品陳列'
                            ,[FIELD26] AS '午班-清潔檢查'
                            ,[FIELD27] AS '午班-設備關閉	'
                            ,[FIELD28] AS '午班-廣告檢查'
                            ,[FIELD29] AS '午班-明日訂單'
                            ,[FIELD29a] AS '午班-明日訂單說明'
                            ,[FIELD30] AS 'FIELD30'
                            ,[FIELD31] AS '午班-清點零用金'
                            ,[FIELD32] AS '午班-額外金額'
                            ,[FIELD33] AS '午班-零用金短／溢	'
                            ,[FIELD34] AS '午班-短溢原因	'
                            ,[FIELD35] AS '營業額'
                            ,[FIELD36] AS 'FIELD36'
                            ,[FIELD37] AS '來客數'
                            ,[FIELD38] AS '單筆均消'
                            ,[FIELD39] AS '清機結帳'
                            ,[FIELD40] AS '交接事項備註'
                            ,[FIELD41] AS '事件記錄'
                            ,[FIELD42] AS 'FIELD42'
                            ,[NAME]
                            FROM [TKMK].[dbo].[TBSTOREDAILY]
                            WHERE [FIELD5]>='{0}' AND [FIELD5]<='{1}'
                            ORDER BY [FIELD5],[FIELD1]
                            ",SDATE,EDATES);

            return SB;

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy/MM/dd"),dateTimePicker2.Value.ToString("yyyy/MM/dd"));
        }

        #endregion
    }
}
