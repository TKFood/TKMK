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
    public partial class frmREPORTSTORES : Form
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

        public frmREPORTSTORES()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            DateTime FirstDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);

            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = LastDay;
        }
        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\門市督導表.frx");

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
                            [ID] AS 'ID'
                            ,[STORE1] AS '門店督導表'
                            ,[STORE2] AS '督導門店'
                            ,[STORE3] AS '督導日期'
                            ,[STORE4] AS '抵達時間'
                            ,[STORE5] AS '離開時間'
                            ,[STORE6] AS '門市督導內容'
                            ,[STORE7] AS '設備設施'
                            ,CONVERT(INT,[STORE8])  AS '評核分數-設備設施'
                            ,[STORE9] AS '門市人員'
                            ,CONVERT(INT,[STORE10])  AS '評核分數-服裝儀容符合規定'
                            ,CONVERT(INT,[STORE11])  AS '評核分數-門店工作日誌確實填寫、記錄及交接'
                            ,CONVERT(INT,[STORE12])  AS '評核分數-正常出勤無臨時請假'
                            ,CONVERT(INT,[STORE13])  AS '評核分數-電話禮儀符合'
                            ,CONVERT(INT,[STORE14])  AS '評核分數-(活動商品、內容)商品賣點抽測'
                            ,[STORE15] AS '抽測內容-(活動商品、內容)商品賣點抽測'
                            ,[STORE16] AS '商品陳列'
                            ,CONVERT(INT,[STORE17])  AS '評核分數-貨架陳列完整無空缺'
                            ,CONVERT(INT,[STORE18])  AS '評核分數-(抽測)架上商品先進先出'
                            ,CONVERT(INT,[STORE19]) AS '評核分數-商品價目牌擺放正確、無毀損'
                            ,[STORE20] AS '未先進先出商品'
                            ,[STORE21] AS '營運狀況 (含活動)'
                            ,CONVERT(INT,[STORE22])  AS '評核分數-無張貼過期POP'
                            ,CONVERT(INT,[STORE23])  AS '評核分數-(活動)POP、展示品、商品擺放正確'
                            ,CONVERT(INT,[STORE24])  AS '評核分數-抽測(活動)商品無缺貨，符合安全庫存量'
                            ,[STORE25] AS '抽測或活動商品當日實際庫存數量'
                            ,CONVERT(INT,[STORE26])  AS '評核分數-找零金金額正確'
                            ,[STORE27] AS '金額-找零金金額正確'
                            ,[STORE28] AS '環境衛生'
                            ,CONVERT(INT,[STORE29])  AS '評核分數-收銀櫃檯無灰塵、個人雜物'
                            ,CONVERT(INT,[STORE30])  AS '評核分數-層架、天花板、騎樓、地面無水漬或蜘蛛網'
                            ,CONVERT(INT,[STORE31])  AS '評核分數-玻璃無指紋'
                            ,CONVERT(INT,[STORE32])  AS '評核分數-商品無直接落地擺放'
                            ,CONVERT(INT,[STORE33])  AS '評核分數-廁所保持清潔、打掃用具歸位整齊擺放'
                            ,[STORE34] AS '庫存管理'
                            ,CONVERT(INT,[STORE35])  AS '評核分數-抽測現場商品庫存量與即時ERP系統一致'
                            ,CONVERT(INT,[STORE36])  AS '評核分數-倉庫排放整齊、(抽測)先進先出'
                            ,[STORE37] AS '庫存管理-抽測商品(抽測現場商品庫存量與即時ERP系統一致)'
                            ,[STORE38] AS '庫存管理-數量(抽測現場商品庫存量與即時ERP系統一致)'
                            ,[STORE39] AS '其他事項'
                            ,[STORE40] AS '輔導狀況-上次督導缺失輔導狀況-已改善'
                            ,[STORE41] AS '輔導狀況-上次督導缺失輔導狀況-未改善說明'
                            ,[STORE42] AS '說明-此次督導缺失、輔導內容及規劃說明'
                            ,[STORE43] AS '其他督導宣達內容'
                            ,[STORE44] AS '門店值班 簽名及回應'
                            ,CONVERT(INT,[STORE8])+CONVERT(INT,[STORE10])+CONVERT(INT,[STORE11])+CONVERT(INT,[STORE12])+CONVERT(INT,[STORE13])+CONVERT(INT,[STORE14])+CONVERT(INT,[STORE17])+CONVERT(INT,[STORE18])+CONVERT(INT,[STORE19])+CONVERT(INT,[STORE22])+CONVERT(INT,[STORE23])+CONVERT(INT,[STORE24])+CONVERT(INT,[STORE26])+CONVERT(INT,[STORE29])+CONVERT(INT,[STORE30])+CONVERT(INT,[STORE31])+CONVERT(INT,[STORE32])+CONVERT(INT,[STORE33])+CONVERT(INT,[STORE35])+CONVERT(INT,[STORE36]) AS '總分'
                            FROM [TKMK].[dbo].[TBSTORESCHECK]
                            WHERE [STORE3]>='{0}' AND [STORE3]<='{1}'
                            ORDER BY [STORE2],[STORE3]
 

                            ", SDATES, EDATES);

            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
        }


        #endregion

    }
}
