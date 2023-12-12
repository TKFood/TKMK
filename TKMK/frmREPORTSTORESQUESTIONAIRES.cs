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
using Google.Apis.Util.Store;

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
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        public frmREPORTSTORESQUESTIONAIRES()
        {
            InitializeComponent();
            comboBox1load();
        }


        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [KINDS],[PARASNAMES],[DVALUES] FROM [TKMK].[dbo].[TBZPARAS] WHERE [KINDS]='frmREPORTSTORESQUESTIONAIRES' ORDER BY DVALUES ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();
     
            dt.Columns.Add("PARASNAMES", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARASNAMES";
            comboBox1.DisplayMember = "PARASNAMES";
            sqlConn.Close();

            //comboBox1.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void SETFASTREPORT(string KINDS,string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQLB1 = new StringBuilder();
            StringBuilder SQLB2 = new StringBuilder();
            StringBuilder SQLB3 = new StringBuilder();
            StringBuilder SQLB4 = new StringBuilder();
            StringBuilder SQLB5 = new StringBuilder();
            StringBuilder SQLB6 = new StringBuilder();
            StringBuilder SQLBC1 = new StringBuilder();
            StringBuilder SQLBC2 = new StringBuilder();
            StringBuilder SQLBC3 = new StringBuilder();
            StringBuilder SQLBC4 = new StringBuilder();
            StringBuilder SQLBC5 = new StringBuilder();
            StringBuilder SQLBC6 = new StringBuilder();


            Report report1 = new Report();         

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

          

            if(KINDS.Equals("門市客群"))
            {
                report1.Load(@"REPORT\門市客群.frx");
                report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
                TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
                SQL1 = SETSQL(SDATES, EDATES);
                table.SelectCommand = SQL1.ToString();
            }
            else if (KINDS.Equals("門市客群購買商品"))
            {
                report1.Load(@"REPORT\門市客群購買商品.frx");
                report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
                TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
                SQLB1 = SETSQLB1(SDATES, EDATES);
                table.SelectCommand = SQLB1.ToString();
                TableDataSource table1= report1.GetDataSource("Table1") as TableDataSource;
                SQLB2 = SETSQLB2(SDATES, EDATES);
                table1.SelectCommand = SQLB2.ToString();
                TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
                SQLB3 = SETSQLB3(SDATES, EDATES);
                table2.SelectCommand = SQLB3.ToString();
                TableDataSource table3 = report1.GetDataSource("Table3") as TableDataSource;
                SQLB4 = SETSQLB4(SDATES, EDATES);
                table3.SelectCommand = SQLB4.ToString();
                TableDataSource table4 = report1.GetDataSource("Table4") as TableDataSource;
                SQLB5 = SETSQLB5(SDATES, EDATES);
                table4.SelectCommand = SQLB5.ToString();
                TableDataSource table5 = report1.GetDataSource("Table5") as TableDataSource;
                SQLB6 = SETSQLB6(SDATES, EDATES);
                table5.SelectCommand = SQLB6.ToString();
            }
            else if (KINDS.Equals("各門市客群購買商品"))
            {
                report1.Load(@"REPORT\各門市客群購買商品.frx");
                report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
                TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
                SQLBC1 = SETSQLC1(SDATES, EDATES);
                table.SelectCommand = SQLBC1.ToString();
                TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
                SQLBC2 = SETSQLC2(SDATES, EDATES);
                table1.SelectCommand = SQLBC2.ToString();
                TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
                SQLBC3 = SETSQLC3(SDATES, EDATES);
                table2.SelectCommand = SQLBC3.ToString();
                TableDataSource table3 = report1.GetDataSource("Table3") as TableDataSource;
                SQLBC4 = SETSQLC4(SDATES, EDATES);
                table3.SelectCommand = SQLBC4.ToString();
                TableDataSource table4 = report1.GetDataSource("Table4") as TableDataSource;
                SQLBC5 = SETSQLC5(SDATES, EDATES);
                table4.SelectCommand = SQLBC5.ToString();
                TableDataSource table5 = report1.GetDataSource("Table5") as TableDataSource;
                SQLBC6 = SETSQLC6(SDATES, EDATES);
                table5.SelectCommand = SQLBC6.ToString();
            }



            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
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

        public StringBuilder SETSQLB1(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT
                            [顧客外觀性別]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY [顧客外觀性別],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0
                            ", SDATES, EDATES);

            return SB;

        }

        public StringBuilder SETSQLB2(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            
                            SELECT
                            [顧客年齡區間]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY [顧客年齡區間],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0

                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLB3(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            
                            SELECT
                            [要送禮還是自己吃]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY [要送禮還是自己吃],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0
                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLB4(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                          
                            SELECT
                            [居住地]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY [居住地],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0

                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLB5(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                          
                            SELECT
                            [本地居住觀光工作]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY [本地居住觀光工作],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0
                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLB6(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT
                            [顧客外觀性別]
                            ,[顧客年齡區間]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY [顧客外觀性別],[顧客年齡區間],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0

                            ", SDATES, EDATES);

            return SB;

        }

        public StringBuilder SETSQLC1(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT
                            [門市]
                            ,[顧客外觀性別]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY  [門市],[顧客外觀性別],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0
                            ", SDATES, EDATES);

            return SB;

        }

        public StringBuilder SETSQLC2(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                             
                            SELECT
                            [門市]
                            ,[顧客年齡區間]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY  [門市],[顧客年齡區間],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0

                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLC3(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                             
                            SELECT
                            [門市]
                            ,[要送禮還是自己吃]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY  [門市],[要送禮還是自己吃],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0
                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLC4(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                           
                            SELECT
                            [門市]
                            ,[居住地]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY  [門市],[居住地],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0

                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLC5(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                           
                            SELECT
                            [門市]
                            ,[本地居住觀光工作]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY  [門市],[本地居住觀光工作],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0
                            ", SDATES, EDATES);

            return SB;

        }
        public StringBuilder SETSQLC6(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT
                            [門市]
                            ,[顧客外觀性別]
                            ,[顧客年齡區間]
                            ,POSTB.TB010 品號
                            ,INVMB.MB002 品名
                            ,SUM(POSTB.TB019) 銷售數量
                            ,SUM(POSTB.TB031) 未稅金額
                            FROM [TKMK].[dbo].[TBSTORESQUESTIONAIRES] WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSTB WITH(NOLOCK) ON TB008=[發票號碼]
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            WHERE 1=1
                            AND ISNULL([發票號碼],'')<>''
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND CONVERT(NVARCHAR,[時間戳記],112)>='{0}' AND CONVERT(NVARCHAR,[時間戳記],112)<='{1}' 
                            GROUP BY  [門市],[顧客外觀性別],[顧客年齡區間],POSTB.TB010,INVMB.MB002
                            HAVING SUM(POSTB.TB031)>0

                            ", SDATES, EDATES);

            return SB;

        }

        public void GET_GOOGLESHEETS()
        {
            string spreadsheetId = "1pORwOtwkaeife1lYFI7yuiT2jYMr1UCXr6FtwzU4WQE";
            string range = "表單回應 1!A1:C10"; // 修改为您的表格和范围
            string credentialsPath = "C:/A1_Github/TKMK/TKMK/LICENSES/tkfood-2023-711945ea86c4.json";

            if (!File.Exists(credentialsPath))
            {
                MessageBox.Show("credentialsPath not exists");
                return;
            }
            else
            {
                //GoogleCredential credential;
                //using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
                //{
                //    credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
                //}

                var credential = GoogleCredential.FromFile(credentialsPath).CreateScoped(SheetsService.Scope.Spreadsheets);

                // Google Sheets API
                var service = GetSheetsService(credentialsPath);

                // 从Google Sheets获取数据
                var data = GetGoogleSheetsData(service, spreadsheetId, range); 
            }
            

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
            SpreadsheetsResource.ValuesResource.GetRequest request =  sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

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
            SETFASTREPORT(comboBox1.Text,dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GET_GOOGLESHEETS();
        }
        #endregion


    }
}
