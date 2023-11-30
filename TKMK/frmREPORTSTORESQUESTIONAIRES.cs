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
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;


namespace TKMK
{
    public partial class frmREPORTSTORESQUESTIONAIRES : Form
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


        public frmREPORTSTORESQUESTIONAIRES()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\門市客群.frx");

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
                            SELECT 
                             [ID]
                            ,[時間戳記]
                            ,[門市]
                            ,[填寫人]
                            ,[有沒有購買商品]
                            ,[發票號碼]
                            ,[顧客外觀性別]
                            ,[顧客年齡區間]
                            ,[要送禮還是自己吃]
                            ,[居住地]
                            ,[本地居住觀光工作]
                            ,[職業或行業]
                            ,[了解到最新消息動態]
                            ,[是否有成為老楊的會員]
                            ,[沒有成為老楊的會員的原因]
                            ,[打算去嘉義哪裡走走]
                            ,[打算去嘉義哪裡走走-其他]
                            ,[其他記錄]
                            ,1 AS COUNTS
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES]
                            WHERE CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}'

                            ", SDATES, EDATES);

            return SB;

        }

        public void GET_GOOGLESHEETS()
        {
            string spreadsheetId = "1pORwOtwkaeife1lYFI7yuiT2jYMr1UCXr6FtwzU4WQE";
            string range = "Sheet1!A1:C10"; // 修改为您的表格和范围
            string credentialsPath = "C:/A1_Github/TKMK/TKMK/LICENSES/client_secret_126586316141-62di5sr2lu7s6lfc96d3ul4k61al0s0c.apps.googleusercontent.com.json";

            if (!File.Exists(credentialsPath))
            {
                MessageBox.Show("credentialsPath not exists");
                return;
            }

            // Google Sheets API
            var service = GetSheetsService(credentialsPath);         
          
            // 从Google Sheets获取数据
            var data = GetGoogleSheetsData(service, spreadsheetId, range);

        }

        static SheetsService GetSheetsService(string credentialsPath)
        {
            var credential = GoogleCredential.FromFile(credentialsPath)
                .CreateScoped(SheetsService.Scope.Spreadsheets);

            var sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google Sheets to Database"
            });

            return sheetsService;
        }

        static List<List<string>> GetGoogleSheetsData(SheetsService sheetsService, string spreadsheetId, string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
                sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;

            var data = new List<List<string>>();

            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    data.Add(row.Select(cell => cell.ToString()).ToList());
                }
            }

            return data;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GET_GOOGLESHEETS();
        }
        #endregion


    }
}
