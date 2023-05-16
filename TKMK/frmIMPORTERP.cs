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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.OleDb;
using System.Net;

namespace TKMK
{
    public partial class frmIMPORTERP : Form
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
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;

        string _filename = null;
        string _path = null;
        DataTable EXCEL = null;


        public frmIMPORTERP()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            textBox1.Text = "106604";

            SET_TEXTBOX();

        }

        #region FUNCTION

        public void SET_TEXTBOX()
        {
            string MESS = "";

            MESS = MESS + "1:本程式限在POS機上執行"+Environment.NewLine;
            MESS = MESS + "2:本程式需先下載 銷售明細表 xls格式，再匯入暫存DB中，並整理出POST相關資料" + Environment.NewLine;
            MESS = MESS + "3:本程式由暫存DB，再匯到POS的本機DB" + Environment.NewLine;
            MESS = MESS + "4:用POS的上傳功能，將本機DB資料匯入到ERP的DB中" + Environment.NewLine;

            textBox2.Text = MESS;
        }
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
            Sequel.AppendFormat(@"SELECT [PARASNAMES],[DVALUES] FROM [TKMK].[dbo].[TBZPARAS] WHERE [PARASNAMES] IN ('是否匯入Y','是否匯入N')  ORDER BY [DVALUES] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("DVALUES", typeof(string));      

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "DVALUES";
            comboBox1.DisplayMember = "DVALUES";
            sqlConn.Close();


        }

        public void comboBox2load()
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
            Sequel.AppendFormat(@"SELECT [PARASNAMES],[DVALUES] FROM [TKMK].[dbo].[TBZPARAS] WHERE KINDS IN ('TBJabezPOSSTORES')  ORDER BY PARASNAMES ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARASNAMES", typeof(string));
            dt.Columns.Add("DVALUES", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "DVALUES";
            comboBox2.DisplayMember = "PARASNAMES";
            sqlConn.Close();


        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(comboBox2.Text))
            {
                textBox1.Text = comboBox2.SelectedValue.ToString();
            }
        }
        public void Search(string YYYYMM,string ISIMPORT)
        {
            DataSet ds = new DataSet();
            StringBuilder sbSqlQUERY = new StringBuilder();


            try
            {
                sbSql.Clear();
                sbSqlQUERY.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                if (ISIMPORT.Equals("Y"))
                {
                    sbSqlQUERY.AppendFormat(@" 
                                            AND  REPLACE(ISNULL(TA001,'')+ISNULL(TA002,'')+ISNULL(TA003,'')+ISNULL(TA006,''),' ','')  IN (SELECT  REPLACE(TA001+TA002+TA003+TA006,' ','') FROM [TK].dbo.POSTA WHERE TA002 IN ('106702')  AND TA003 IN ('900') )
                                            ");
                }
                else
                {
                    sbSqlQUERY.AppendFormat(@" 
                                            AND  REPLACE(ISNULL(TA001,'')+ISNULL(TA002,'')+ISNULL(TA003,'')+ISNULL(TA006,''),' ','') NOT IN (SELECT  REPLACE(TA001+TA002+TA003+TA006,' ','') FROM [TK].dbo.POSTA WHERE TA002 IN ('106702')  AND TA003 IN ('900') )
                                            ");
                }

                sbSql.AppendFormat(@"
                                    SELECT
                                    [營業點]
                                    ,[機台]
                                    ,[日期]
                                    ,[序號]
                                    ,[自訂序號]
                                    ,[時間]
                                    ,[訂單屬性]
                                    ,[發票]
                                    ,[統編]
                                    ,[收銀員]
                                    ,[會員]
                                    ,[註記]
                                    ,[附餐內容物]
                                    ,[商品編號]
                                    ,[商品名稱]
                                    ,[單價]
                                    ,[數量]
                                    ,[小計]
                                    ,[口味]
                                    ,[加料]
                                    ,[容量]
                                    ,[總金額]
                                    ,[總折扣]
                                    ,[明細折扣]
                                    ,[明細金額]
                                    ,[TA001]
                                    ,[TA002]
                                    ,[TA003]
                                    ,[TA006]
                                    ,[TB007]
                                    ,[TC007]
                                     
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE [日期] LIKE '%{0}%'
                                    {1}
                                    ORDER BY [營業點],[機台],[日期],[序號],[時間],[自訂序號]

                                         ", YYYYMM, sbSqlQUERY.ToString());

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                }
                else
                {
                    dataGridView1.DataSource = ds.Tables["ds"];
                    dataGridView1.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;                       

                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                }



            }
            catch
            {

            }
            finally
            {

            }
        }
        public void CHECKADDDATA()
        {
            //新增暫存資料   
            //並給明細序號
            //計算折扣後金額  
            IMPORTEXCEL();
            //ImportCSV();



            //檢查暫存中的新資料
            DataTable NEWTDATATABLE = SEARCHNEWDATA();
            //匯入到TBJabezPOS中
            if (NEWTDATATABLE != null && NEWTDATATABLE.Rows.Count > 0)
            {
                ADD_TO_TBJabezPOS(NEWTDATATABLE);
            }


        }

        public DataTable SEARCHTBJabezPOS()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            //THISYEARS = "21";

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                //核準過TASK_RESULT='0'
                //AND DOC_NBR  LIKE 'QC1002{0}%'

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [機台]
                                    ,[日期]
                                    ,[序號]
                                    ,[時間]
                                    ,[商品編號]
                                    ,[商品規格]
                                    ,[單價]
                                    ,[數量]
                                    ,[小計]
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void IMPORTEXCEL()
        {
            //記錄選到的檔案路徑
            _filename = null;
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {

            }
            if (dr == DialogResult.Cancel)
            {

            }


            textBox3.Text = od.FileName.ToString();
            _filename = od.FileName.ToString();

            //string pathToExcelFile = @"F:\rcPOS9002_20230515.xls";
            string pathToExcelFile = _filename;
            string sheetName = "";
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathToExcelFile + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            DataTable dataTabl = new DataTable();
            DataTable dataTab2 = new DataTable();
            DataTable dataTab3 = new DataTable();

            //自訂dataTab2的匯入欄位+轉decimal
            dataTab2.Columns.Add("營業點", typeof(string));
            dataTab2.Columns.Add("機台", typeof(string));
            dataTab2.Columns.Add("日期", typeof(string));
            dataTab2.Columns.Add("序號", typeof(string));
            dataTab2.Columns.Add("時間", typeof(string));
            dataTab2.Columns.Add("訂單屬性", typeof(string));
            dataTab2.Columns.Add("發票", typeof(string));
            dataTab2.Columns.Add("統編", typeof(string));
            dataTab2.Columns.Add("收銀員", typeof(string));
            dataTab2.Columns.Add("會員", typeof(string));
            dataTab2.Columns.Add("註記", typeof(string));
            dataTab2.Columns.Add("附餐/內容物", typeof(string));
            dataTab2.Columns.Add("商品編號", typeof(string));
            dataTab2.Columns.Add("商品名稱", typeof(string));
            dataTab2.Columns.Add("單價", typeof(decimal));
            dataTab2.Columns.Add("數量", typeof(decimal));
            dataTab2.Columns.Add("小計", typeof(decimal));
            dataTab2.Columns.Add("口味", typeof(string));
            dataTab2.Columns.Add("加料", typeof(string));
            dataTab2.Columns.Add("容量", typeof(string));

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                dataTabl = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dataTabl != null)
                {
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        sheetName = sheetName + "$";
                    }
                    else
                    {
                        sheetName = dataTabl.Rows[0]["TABLE_NAME"].ToString();
                    }
                }
                string sql = string.Format("SELECT * FROM [{0}]", sheetName);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, connection))
                {
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);
                    dataTab2 = dataSet.Tables[0];
                    // 讀取到的資料存放在 dataTable 變數中
                }
            }

            dataTab3 = dataTab2;

            //另外新增自訂序號
            dataTab3.Columns.Add("自訂序號", typeof(int));
            // 建立一個 DataView，並將 DataTable 設為資料來源
            DataView dv = new DataView(dataTab3);
            // 設定排序條件，先按 Name 欄位進行升序排序，如果 Name 相同，再按 Age 欄位進行升序排序，如果 Age 相同，最後按 Salary 欄位進行升序排序
            dv.Sort = "營業點 ASC, 機台 ASC, 日期 ASC,序號 ASC,商品編號 ASC";
            // 將排序後的 DataView 轉換回 DataTable
            DataTable sortedDt = dv.ToTable();

            DataTable KEYDT = SET_DETAILS_KEYS(sortedDt);

            ADD_TO_TBJabezPOS_TEMP(KEYDT);
            UPDATE_TBJabezPOS_TEMP();
        }

        //處理明細序號
        //處理金額是空白

        public DataTable SET_DETAILS_KEYS(DataTable SDT)
        {
            int ROWS = 0;
            int KEYORDER = 1;
                       
            string 營業點 = "";
            string 機台 = "";
            string 日期 = "";
            string 序號 = "";
            string NEW營業點 = "";
            string NEW機台 = "";
            string NEW日期 = "";
            string NEW序號 = "";

            foreach (DataRow DR in SDT.Rows)
            {
                NEW營業點 = DR["營業點"].ToString();
                NEW機台 = DR["機台"].ToString();
                NEW日期 = DR["日期"].ToString();
                NEW序號 = DR["序號"].ToString();

                if(營業點.Equals(NEW營業點) && 機台.Equals(NEW機台) && 日期.Equals(NEW日期) && 序號.Equals(NEW序號)  )
                {
                    KEYORDER = KEYORDER + 1;
                    SDT.Rows[ROWS]["自訂序號"] = KEYORDER;                    
                }
                else
                {
                    KEYORDER = 1;
                    SDT.Rows[ROWS]["自訂序號"] = KEYORDER;                   
                }

                if(string.IsNullOrEmpty(DR["單價"].ToString()))
                {
                    SDT.Rows[ROWS]["單價"] = 0;
                }
                if (string.IsNullOrEmpty(DR["數量"].ToString()))
                {
                    SDT.Rows[ROWS]["數量"] = 0;
                }


                營業點 = NEW營業點;
                機台 = NEW機台;
                日期 = NEW日期;
                序號 = NEW序號;

                ROWS = ROWS + 1;
            }

            return SDT;
        }

        public void ImportCSV()
        {
            //記錄選到的檔案路徑
            _filename = null;
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {

            }
            if (dr == DialogResult.Cancel)
            {

            }


            textBox3.Text = od.FileName.ToString();
            _filename= od.FileName.ToString();


            var table = new DataTable();
            using (var streamReader = new StreamReader(_filename))
            {
                // 設定分隔符號為逗號
                var separator = ',';
                // 讀取 CSV 標題行
                var header = streamReader.ReadLine();
                if (!string.IsNullOrEmpty(header))
                {
                    var columns = header.Split(separator);
                    foreach (var column in columns)
                    {
                        table.Columns.Add(column.Trim());
                    }
                }

                // 讀取 CSV 內容
                while (!streamReader.EndOfStream)
                {
                    var line = streamReader.ReadLine();
                    if (!string.IsNullOrEmpty(line))
                    {
                        var values = line.Split(separator);
                        var row = table.NewRow();
                        for (int i = 0; i < values.Length; i++)
                        {
                            row[i] = values[i].Trim();
                        }
                        table.Rows.Add(row);
                    }
                }
            }

            //return table;

            ADD_TO_TBJabezPOS_TEMP(table);
           

        }

    
        /// <summary>
        ///新增暫存資料到 TBJabezPOS_TEMP 
        /// </summary>
        /// <param name="DT"></param>
        public void ADD_TO_TBJabezPOS_TEMP(DataTable DT)
        {
            CLEAR_TBJabezPOS_TEMP();
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);           
           
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
            {
                sqlConn.Open();
                bulkCopy.DestinationTableName = "TBJabezPOS_TEMP";

                //對應資料行
                //   bulkCopy.ColumnMappings.Add("來源欄位", "目標欄位");
                bulkCopy.ColumnMappings.Add("營業點", "營業點");
                bulkCopy.ColumnMappings.Add("機台", "機台");
                bulkCopy.ColumnMappings.Add("日期", "日期");
                bulkCopy.ColumnMappings.Add("序號", "序號");
                bulkCopy.ColumnMappings.Add("自訂序號", "自訂序號");
                bulkCopy.ColumnMappings.Add("時間", "時間");
                bulkCopy.ColumnMappings.Add("訂單屬性", "訂單屬性");
                bulkCopy.ColumnMappings.Add("發票", "發票");
                bulkCopy.ColumnMappings.Add("統編", "統編");
                bulkCopy.ColumnMappings.Add("收銀員", "收銀員");
                bulkCopy.ColumnMappings.Add("會員", "會員");
                bulkCopy.ColumnMappings.Add("註記", "註記");
                bulkCopy.ColumnMappings.Add("附餐/內容物", "附餐內容物");
                bulkCopy.ColumnMappings.Add("商品編號", "商品編號");
                bulkCopy.ColumnMappings.Add("商品名稱", "商品名稱");
                bulkCopy.ColumnMappings.Add("單價", "單價");
                bulkCopy.ColumnMappings.Add("數量", "數量");
                bulkCopy.ColumnMappings.Add("小計", "小計");
                bulkCopy.ColumnMappings.Add("口味", "口味");
                bulkCopy.ColumnMappings.Add("加料", "加料");
                bulkCopy.ColumnMappings.Add("容量", "容量");

                try
                {
                    bulkCopy.WriteToServer(DT);
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                }
            }
        }

        public void UPDATE_TBJabezPOS_TEMP()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                //更新 總金額
                //更新 總折扣<0 的合計
                //更新 明細折扣 分攤
                //更新 明細折扣的最後一筆=0
                //更新 明細折扣的最後一筆=該筆內的其他明細折扣合計
                sbSql.AppendFormat(@" 
                               
                                    UPDATE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    SET [TBJabezPOS_TEMP].[總金額]=TEMP.TMONEYS
                                    FROM (
                                    SELECT 
                                     [營業點]
                                    ,[機台]
                                    ,[日期]
                                    ,[序號]
                                    ,SUM([小計]) AS TMONEYS
                                    FROM [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE [小計]>=0
                                    GROUP BY  [營業點],[機台],[日期],[序號]
                                    ) AS TEMP
                                    WHERE TEMP.[營業點]=[TBJabezPOS_TEMP].[營業點] AND TEMP.[機台]=[TBJabezPOS_TEMP].[機台] AND TEMP.[日期]=[TBJabezPOS_TEMP].[日期] AND TEMP.[序號]=[TBJabezPOS_TEMP].[序號]
                                    AND [TBJabezPOS_TEMP].[總金額]<>TEMP.TMONEYS
                                    AND [TBJabezPOS_TEMP].[小計]>0
                 
                                    UPDATE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    SET [TBJabezPOS_TEMP].[總折扣]=TEMP.[小計]
                                    FROM (
                                    SELECT  
                                    [營業點]
                                    ,[機台]
                                    ,[日期]
                                    ,[序號]
                                    ,[自訂序號]
                                    ,[小計]
                                    FROM [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE [商品名稱]='小計後加減價'
                                    ) AS TEMP
                                    WHERE TEMP.[營業點]=[TBJabezPOS_TEMP].[營業點] AND TEMP.[機台]=[TBJabezPOS_TEMP].[機台] AND TEMP.[日期]=[TBJabezPOS_TEMP].[日期] AND TEMP.[序號]=[TBJabezPOS_TEMP].[序號]
                                    AND [TBJabezPOS_TEMP].總折扣<>TEMP.[小計] 
  
                                     UPDATE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                     SET [明細折扣]=ROUND([小計]/[總金額]*[總折扣],0)
                                     WHERE [總折扣]<>0 AND [小計]>0

                                     UPDATE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                     SET [明細折扣]=0
                                     FROM 
                                     (
                                     SELECT 
                                     [營業點]
                                    ,[機台]
                                    ,[日期]
                                    ,[序號]
                                    ,MAX([自訂序號]) [自訂序號]
                                    FROM [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE [明細折扣]<0
                                    GROUP BY  [營業點],[機台],[日期],[序號]
                                    ) AS TEMP
                                    WHERE TEMP.[營業點]=[TBJabezPOS_TEMP].[營業點] AND TEMP.[機台]=[TBJabezPOS_TEMP].[機台] AND TEMP.[日期]=[TBJabezPOS_TEMP].[日期] AND TEMP.[序號]=[TBJabezPOS_TEMP].[序號]
                                    AND TEMP.[自訂序號]=[TBJabezPOS_TEMP].[自訂序號]

                                    UPDATE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    SET [TBJabezPOS_TEMP].[明細折扣]=[TBJabezPOS_TEMP].[總折扣]-DISTMONEYS
                                    FROM 
                                    (
                                    SELECT 
                                     [營業點]
                                    ,[機台]
                                    ,[日期]
                                    ,[序號]
                                    ,MAX([自訂序號]) [自訂序號]
                                    ,(SELECT SUM([明細折扣]) FROM [TKMK].[dbo].[TBJabezPOS_TEMP] TEMP1 WHERE TEMP1.營業點=[TBJabezPOS_TEMP].營業點 AND  TEMP1.[機台]=[TBJabezPOS_TEMP].[機台] AND  TEMP1.[日期]=[TBJabezPOS_TEMP].[日期] AND  TEMP1.[序號]=[TBJabezPOS_TEMP].[序號]   ) AS DISTMONEYS
                                    FROM [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE [總折扣]<0 AND [明細折扣]=0
                                    GROUP BY  [營業點],[機台],[日期],[序號]
                                    ) AS TEMP
                                    WHERE TEMP.[營業點]=[TBJabezPOS_TEMP].[營業點] AND TEMP.[機台]=[TBJabezPOS_TEMP].[機台] AND TEMP.[日期]=[TBJabezPOS_TEMP].[日期] AND TEMP.[序號]=[TBJabezPOS_TEMP].[序號]
                                    AND TEMP.[自訂序號]=[TBJabezPOS_TEMP].[自訂序號]
                                    AND [小計]>0
                                    
                                    DELETE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE [商品名稱]='小計後加減價'

                                    UPDATE [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    SET [明細金額]=[小計]+[明細折扣]
                                        ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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
        /// <summary>
        /// 清空暫存的 TBJabezPOS_TEMP
        /// </summary>
        public void CLEAR_TBJabezPOS_TEMP()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                                
                sbSql.AppendFormat(@" 
                                    DELETE  [TKMK].[dbo].[TBJabezPOS_TEMP]
                                        ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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

        public DataTable SEARCHNEWDATA()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            //THISYEARS = "21";

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                //核準過TASK_RESULT='0'
                //AND DOC_NBR  LIKE 'QC1002{0}%'

                sbSql.AppendFormat(@"                                   
                                    SELECT 
                                    *
                                    FROM [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE REPLACE([營業點]+[機台]+[日期]+[序號],' ','') NOT IN (SELECT REPLACE([營業點]+[機台]+[日期]+[序號],' ','') FROM [TKMK].[dbo].[TBJabezPOS])  
                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        /// <summary>
        /// 新增到TBJabezPOS
        /// </summary>
        /// <param name="DT"></param>
        public void ADD_TO_TBJabezPOS(DataTable DT)
        {
            CLEAR_TBJabezPOS_TEMP();
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
            {
                sqlConn.Open();
                bulkCopy.DestinationTableName = "TBJabezPOS";
                //對應資料行
                //   bulkCopy.ColumnMappings.Add("來源欄位", "目標欄位");
                bulkCopy.ColumnMappings.Add("營業點", "營業點");
                bulkCopy.ColumnMappings.Add("機台", "機台");
                bulkCopy.ColumnMappings.Add("日期", "日期");
                bulkCopy.ColumnMappings.Add("序號", "序號");
                bulkCopy.ColumnMappings.Add("自訂序號", "自訂序號");
                bulkCopy.ColumnMappings.Add("時間", "時間");
                bulkCopy.ColumnMappings.Add("訂單屬性", "訂單屬性");
                bulkCopy.ColumnMappings.Add("發票", "發票");
                bulkCopy.ColumnMappings.Add("統編", "統編");
                bulkCopy.ColumnMappings.Add("收銀員", "收銀員");
                bulkCopy.ColumnMappings.Add("會員", "會員");
                bulkCopy.ColumnMappings.Add("註記", "註記");
                bulkCopy.ColumnMappings.Add("附餐內容物", "附餐內容物");
                bulkCopy.ColumnMappings.Add("商品編號", "商品編號");
                bulkCopy.ColumnMappings.Add("商品名稱", "商品名稱");
                bulkCopy.ColumnMappings.Add("單價", "單價");
                bulkCopy.ColumnMappings.Add("數量", "數量");
                bulkCopy.ColumnMappings.Add("小計", "小計");
                bulkCopy.ColumnMappings.Add("口味", "口味");
                bulkCopy.ColumnMappings.Add("加料", "加料");
                bulkCopy.ColumnMappings.Add("容量", "容量");
                bulkCopy.ColumnMappings.Add("總金額", "總金額");
                bulkCopy.ColumnMappings.Add("總折扣", "總折扣");
                bulkCopy.ColumnMappings.Add("明細折扣", "明細折扣");
                bulkCopy.ColumnMappings.Add("明細金額", "明細金額");

                try
                {
                    bulkCopy.WriteToServer(DT);
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                }
            }
        }

        public void UPDATE_TBJabezPOS_TA001TA002TA003(string TA002)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    
                                    UPDATE [TKMK].[dbo].[TBJabezPOS]
                                    SET [TA001]=[日期]
                                    WHERE ISNULL([TA001],'')=''

                                    UPDATE [TKMK].[dbo].[TBJabezPOS]
                                    SET [TA002]='{0}'
                                    WHERE ISNULL([TA002],'')=''
                                 

                                    UPDATE [TKMK].[dbo].[TBJabezPOS]
                                    SET [TA003]=[DVALUES]
                                    FROM [TKMK].[dbo].[TBZPARAS]
                                    WHERE ISNULL([TA003],'')=''
                                    AND [TBZPARAS].KINDS='TBJabezPOS' AND [TBZPARAS].PARASNAMES='機號'
                                        ", TA002);


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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

        public void UPDATE_TBJabezPOS_TA006()
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                
                sbSql.AppendFormat(@"                                    
                                    SELECT  
                                    [TA001],[TA002],[TA003]
                                    ,RIGHT('00000'+CAST(row_number() OVER(PARTITION BY [TA001],[TA002],[TA003] ORDER BY [TA001],[TA002],[TA003]) AS nvarchar(10)),5)  AS SEQ 
                                    ,([營業點]+[機台]+[日期]+[序號]) AS 'KEYS'
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE ISNULL(TA006,'')=''
                                    GROUP BY [TA001],[TA002],[TA003],[營業點]+[機台]+[日期]+[序號]
                                    ORDER BY [營業點]+[機台]+[日期]+[序號]

                                         ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    UPDATE_TBJabezPOS_TA006_ACT(ds.Tables["ds"]);
                }




            }
            catch
            {

            }
            finally
            {

            }
        }

        public void UPDATE_TBJabezPOS_TA006_ACT(DataTable DT)
        {
            StringBuilder SQLEXECUT = new StringBuilder();

            foreach(DataRow DR in DT.Rows)
            {
                SQLEXECUT.AppendFormat(@"
                                        UPDATE [TKMK].[dbo].[TBJabezPOS]
                                        SET [TA006]='{1}'
                                        WHERE ([營業點]+[機台]+[日期]+[序號])='{0}'

                                        ", DR["KEYS"].ToString(), DR["SEQ"].ToString());
            }


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql = SQLEXECUT;
                //sbSql.AppendFormat(@"  ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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

        public void UPDATE_TBJabezPOS_TB007TC007()
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.AppendFormat(@"                                    
                                   SELECT  
                                   *
                                    ,RIGHT('0000'+CAST(row_number() OVER(PARTITION BY [TA001],[TA002],[TA003],[TA006] ORDER BY [TA001],[TA002],[TA003],[TA006]) AS nvarchar(10)),4)  AS SEQ 
                                    ,([TA001]+[TA002]+[TA003]+[TA006]+CONVERT(NVARCHAR,[自訂序號])) AS 'KEYS'
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE ISNULL([TB007],'')=''


                                         ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    UPDATE_TBJabezPOS_TB007TC007_ACT(ds.Tables["ds"]);
                }




            }
            catch
            {

            }
            finally
            {

            }
        }

        public void UPDATE_TBJabezPOS_TB007TC007_ACT(DataTable DT)
        {
            StringBuilder SQLEXECUT = new StringBuilder();

            foreach (DataRow DR in DT.Rows)
            {
                SQLEXECUT.AppendFormat(@"
                                        UPDATE [TKMK].[dbo].[TBJabezPOS]
                                        SET [TB007]='{1}'
                                        WHERE ([TA001]+[TA002]+[TA003]+[TA006]+CONVERT(NVARCHAR,[自訂序號]))='{0}'
                                        
                                        UPDATE [TKMK].[dbo].[TBJabezPOS]
                                        SET [TC007]='001'

                                        ", DR["KEYS"].ToString(), DR["SEQ"].ToString());
            }


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql = SQLEXECUT;
                //sbSql.AppendFormat(@"  ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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

        public void ADD_ERP_POSTAPOSTBPOSTC()
        {
            
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbLOCAL"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

             
                sbSql.AppendFormat(@"  

                                    INSERT INTO [COSMOS_POS].dbo.POSTA
                                    (
                                    [COMPANY]
                                    ,[CREATOR]
                                    ,[USR_GROUP]
                                    ,[CREATE_DATE]
                                    ,[MODIFIER]
                                    ,[MODI_DATE]
                                    ,[FLAG]
                                    ,[CREATE_TIME]
                                    ,[MODI_TIME]
                                    ,[TRANS_TYPE]
                                    ,[TRANS_NAME]
                                    ,[sync_date]
                                    ,[sync_time]
                                    ,[sync_mark]
                                    ,[sync_count]
                                    ,[DataUser]
                                    ,[DataGroup]
                                    ,[TA001]
                                    ,[TA002]
                                    ,[TA003]
                                    ,[TA004]
                                    ,[TA005]
                                    ,[TA006]
                                    ,[TA007]
                                    ,[TA008]
                                    ,[TA009]
                                    ,[TA010]
                                    ,[TA011]
                                    ,[TA012]
                                    ,[TA013]
                                    ,[TA014]
                                    ,[TA015]
                                    ,[TA016]
                                    ,[TA017]
                                    ,[TA018]
                                    ,[TA019]
                                    ,[TA020]
                                    ,[TA021]
                                    ,[TA022]
                                    ,[TA023]
                                    ,[TA024]
                                    ,[TA025]
                                    ,[TA026]
                                    ,[TA027]
                                    ,[TA028]
                                    ,[TA029]
                                    ,[TA030]
                                    ,[TA031]
                                    ,[TA032]
                                    ,[TA033]
                                    ,[TA034]
                                    ,[TA035]
                                    ,[TA036]
                                    ,[TA037]
                                    ,[TA038]
                                    ,[TA039]
                                    ,[TA040]
                                    ,[TA041]
                                    ,[TA042]
                                    ,[TA043]
                                    ,[TA044]
                                    ,[TA045]
                                    ,[TA046]
                                    ,[TA047]
                                    ,[TA048]
                                    ,[TA049]
                                    ,[TA050]
                                    ,[TA051]
                                    ,[TA052]
                                    ,[TA053]
                                    ,[TA054]
                                    ,[TA055]
                                    ,[TA056]
                                    ,[TA057]
                                    ,[TA058]
                                    ,[TA059]
                                    ,[TA060]
                                    ,[TA061]
                                    ,[TA062]
                                    ,[TA063]
                                    ,[TA064]
                                    ,[TA065]
                                    ,[TA066]
                                    ,[TA067]
                                    ,[TA068]
                                    ,[TA069]
                                    ,[TA070]
                                    ,[TA071]
                                    ,[TA072]
                                    ,[TA073]
                                    ,[TA074]
                                    ,[TA075]
                                    ,[TA076]
                                    ,[TA077]
                                    ,[TA078]
                                    ,[TA079]
                                    ,[TA080]
                                    ,[TA081]
                                    ,[TA082]
                                    ,[TA083]
                                    ,[TA084]
                                    ,[TA085]
                                    ,[TA086]
                                    ,[TA087]
                                    ,[TA088]
                                    ,[TA089]
                                    ,[TA090]
                                    ,[TA091]
                                    ,[TA092]
                                    ,[TA093]
                                    ,[TA094]
                                    ,[TA095]
                                    ,[TA096]
                                    ,[TA097]
                                    ,[TA098]
                                    ,[TA099]
                                    ,[TA100]
                                    ,[TA101]
                                    ,[TA102]
                                    ,[TA103]
                                    ,[TA104]
                                    ,[TA105]
                                    ,[TA106]
                                    ,[TA107]
                                    ,[TA108]
                                    ,[TA109]
                                    ,[TA110]
                                    ,[TA111]
                                    ,[TA112]
                                    ,[TA113]
                                    ,[TA114]
                                    ,[TA115]
                                    ,[TA116]
                                    ,[TA117]
                                    ,[UDF01]
                                    ,[UDF02]
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                    ,[UDF06]
                                    ,[UDF07]
                                    ,[UDF08]
                                    ,[UDF09]
                                    ,[UDF10]
                                    )

                                    SELECT
                                    'TK' [COMPANY]
                                    ,'150073' [CREATOR]
                                    ,'' [USR_GROUP]
                                    ,[TA001] [CREATE_DATE]
                                    ,'' [MODIFIER]
                                    ,'' [MODI_DATE]
                                    ,0 [FLAG]
                                    ,[時間] [CREATE_TIME]
                                    ,'' [MODI_TIME]
                                    ,'' [TRANS_TYPE]
                                    ,'' [TRANS_NAME]
                                    ,'' [sync_date]
                                    ,'' [sync_time]
                                    ,'N' [sync_mark]
                                    ,0 [sync_count]
                                    ,'' [DataUser]
                                    ,'' [DataGroup]
                                    ,[TA001] [TA001]
                                    ,[TA002] [TA002]
                                    ,[TA003] [TA003]
                                    ,[TA001] [TA004]
                                    ,[時間] [TA005]
                                    ,[TA006] [TA006]
                                    ,'150073' [TA007]
                                    ,'' [TA008]
                                    ,'' [TA009]
                                    ,'' [TA010]
                                    ,1 [TA011]
                                    ,'' [TA012]
                                    ,'' [TA013]
                                    ,[發票] [TA014]
                                    ,0 [TA015]
                                    ,SUM([數量]) [TA016]
                                    ,SUM([明細金額]) [TA017]
                                    ,0 [TA018]
                                    ,SUM([明細金額]) [TA019]
                                    ,0 [TA020]
                                    ,CONVERT(INT,SUM([明細金額])*0.05) [TA021]
                                    ,0 [TA022]
                                    ,0 [TA023]
                                    ,SUM([明細金額]) [TA024]
                                    ,SUM([明細金額]) [TA025]
                                    ,(SUM([明細金額])-CONVERT(INT,SUM([明細金額])*0.05)) [TA026]
                                    ,CONVERT(INT,SUM([明細金額])*0.05) [TA027]
                                    ,0 [TA028]
                                    ,0 [TA029]
                                    ,0 [TA030]
                                    ,0 [TA031]
                                    ,0 [TA032]
                                    ,0 [TA033]
                                    ,0 [TA034]
                                    ,'' [TA035]
                                    ,'0' [TA036]
                                    ,'1' [TA037]
                                    ,'1' [TA038]
                                    ,'' [TA039]
                                    ,'N' [TA040]
                                    ,[發票] [TA041]
                                    ,'' [TA042]
                                    ,'' [TA043]
                                    ,'' [TA044]
                                    ,'' [TA045]
                                    , DATEPART(DW,[TA001])   [TA046]
                                    ,'' [TA047]
                                    ,'' [TA048]
                                    ,'' [TA049]
                                    ,'' [TA050]
                                    ,0 [TA051]
                                    ,0 [TA052]
                                    ,'N' [TA053]
                                    ,'' [TA054]
                                    ,'' [TA055]
                                    ,'' [TA056]
                                    ,'' [TA057]
                                    ,'' [TA058]
                                    ,0 [TA059]
                                    ,0 [TA060]
                                    ,0 [TA061]
                                    ,0 [TA062]
                                    ,0 [TA063]
                                    ,0 [TA064]
                                    ,[TA001] [TA065]
                                    ,0 [TA066]
                                    ,'' [TA067]
                                    ,'' [TA068]
                                    ,'' [TA069]
                                    ,0 [TA070]
                                    ,'' [TA071]
                                    ,'' [TA072]
                                    ,'' [TA073]
                                    ,'' [TA074]
                                    ,'' [TA075]
                                    ,'' [TA076]
                                    ,'' [TA077]
                                    ,0 [TA078]
                                    ,'' [TA079]
                                    ,0 [TA080]
                                    ,'' [TA081]
                                    ,'' [TA082]
                                    ,0 [TA083]
                                    ,'' [TA084]
                                    ,0 [TA085]
                                    ,'' [TA086]
                                    ,'' [TA087]
                                    ,'' [TA088]
                                    ,'' [TA089]
                                    ,'N' [TA090]
                                    ,'' [TA091]
                                    ,'' [TA092]
                                    ,'' [TA093]
                                    ,'' [TA094]
                                    ,'' [TA095]
                                    ,'' [TA096]
                                    ,'47730274' [TA097]
                                    ,'47730274' [TA098]
                                    ,'47730274' [TA099]
                                    ,'' [TA100]
                                    ,'' [TA101]
                                    ,0 [TA102]
                                    ,0 [TA103]
                                    ,0 [TA104]
                                    ,0 [TA105]
                                    ,'' [TA106]
                                    ,'N' [TA107]
                                    ,'' [TA108]
                                    ,'' [TA109]
                                    ,'' [TA110]
                                    ,'' [TA111]
                                    ,'' [TA112]
                                    ,'' [TA113]
                                    ,'' [TA114]
                                    ,'' [TA115]
                                    ,'' [TA116]
                                    ,'' [TA117]
                                    ,'' [UDF01]
                                    ,'' [UDF02]
                                    ,'' [UDF03]
                                    ,'' [UDF04]
                                    ,'' [UDF05]
                                    ,0 [UDF06]
                                    ,0 [UDF07]
                                    ,0 [UDF08]
                                    ,0 [UDF09]
                                    ,0 [UDF10]
                                    FROM [192.168.1.105].[TKMK].[dbo].[TBJabezPOS]
                                    WHERE REPLACE([TBJabezPOS].TA001+[TBJabezPOS].TA002+[TBJabezPOS].TA003+[TBJabezPOS].TA006,' ','') NOT IN (SELECT REPLACE(TA001+TA002+TA003+TA006,' ','') FROM [COSMOS_POS].dbo.POSTA  WHERE TA002 IN ('106702')  AND TA003 IN ('900'))
                                    GROUP BY [TA001],[TA002],[TA003],[TA006],[時間],[發票]

                                    INSERT INTO  [COSMOS_POS].dbo.POSTB
                                    (
                                    [COMPANY]
                                    ,[CREATOR]
                                    ,[USR_GROUP]
                                    ,[CREATE_DATE]
                                    ,[MODIFIER]
                                    ,[MODI_DATE]
                                    ,[FLAG]
                                    ,[CREATE_TIME]
                                    ,[MODI_TIME]
                                    ,[TRANS_TYPE]
                                    ,[TRANS_NAME]
                                    ,[sync_date]
                                    ,[sync_time]
                                    ,[sync_mark]
                                    ,[sync_count]
                                    ,[DataUser]
                                    ,[DataGroup]
                                    ,[TB001]
                                    ,[TB002]
                                    ,[TB003]
                                    ,[TB004]
                                    ,[TB005]
                                    ,[TB006]
                                    ,[TB007]
                                    ,[TB008]
                                    ,[TB009]
                                    ,[TB010]
                                    ,[TB011]
                                    ,[TB012]
                                    ,[TB013]
                                    ,[TB014]
                                    ,[TB015]
                                    ,[TB016]
                                    ,[TB017]
                                    ,[TB018]
                                    ,[TB019]
                                    ,[TB020]
                                    ,[TB021]
                                    ,[TB022]
                                    ,[TB023]
                                    ,[TB024]
                                    ,[TB025]
                                    ,[TB026]
                                    ,[TB027]
                                    ,[TB028]
                                    ,[TB029]
                                    ,[TB030]
                                    ,[TB031]
                                    ,[TB032]
                                    ,[TB033]
                                    ,[TB034]
                                    ,[TB035]
                                    ,[TB036]
                                    ,[TB037]
                                    ,[TB038]
                                    ,[TB039]
                                    ,[TB040]
                                    ,[TB041]
                                    ,[TB042]
                                    ,[TB043]
                                    ,[TB044]
                                    ,[TB045]
                                    ,[TB046]
                                    ,[TB047]
                                    ,[TB048]
                                    ,[TB049]
                                    ,[TB050]
                                    ,[TB051]
                                    ,[TB052]
                                    ,[TB053]
                                    ,[TB054]
                                    ,[TB055]
                                    ,[TB056]
                                    ,[TB057]
                                    ,[TB058]
                                    ,[TB059]
                                    ,[TB060]
                                    ,[TB061]
                                    ,[TB062]
                                    ,[TB063]
                                    ,[TB064]
                                    ,[TB065]
                                    ,[TB066]
                                    ,[TB067]
                                    ,[TB068]
                                    ,[TB069]
                                    ,[TB070]
                                    ,[TB071]
                                    ,[TB072]
                                    ,[TB073]
                                    ,[TB074]
                                    ,[TB075]
                                    ,[TB076]
                                    ,[TB077]
                                    ,[TB078]
                                    ,[TB079]
                                    ,[TB080]
                                    ,[TB081]
                                    ,[TB082]
                                    ,[TB083]
                                    ,[TB084]
                                    ,[TB085]
                                    ,[TB086]
                                    ,[TB087]
                                    ,[TB088]
                                    ,[TB089]
                                    ,[TB090]
                                    ,[UDF01]
                                    ,[UDF02]
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                    ,[UDF06]
                                    ,[UDF07]
                                    ,[UDF08]
                                    ,[UDF09]
                                    ,[UDF10]
                                    )

                                   SELECT 
                                    'TK' [COMPANY]
                                    ,'150073' [CREATOR]
                                    ,'' [USR_GROUP]
                                    ,[TA001] [CREATE_DATE]
                                    ,'' [MODIFIER]
                                    ,'' [MODI_DATE]
                                    ,0 [FLAG]
                                    ,[時間] [CREATE_TIME]
                                    ,'' [MODI_TIME]
                                    ,'' [TRANS_TYPE]
                                    ,'' [TRANS_NAME]
                                    ,'' [sync_date]
                                    ,'' [sync_time]
                                    ,'N' [sync_mark]
                                    ,0 [sync_count]
                                    ,'' [DataUser]
                                    ,'' [DataGroup]
                                    ,[TA001] [TB001]
                                    ,[TA002]  [TB002]
                                    ,[TA003] [TB003]
                                    ,[TA001][TB004]
                                    ,[時間]  [TB005]
                                    ,[TA006] [TB006]
                                    ,[TB007] [TB007]
                                    ,'' [TB008]
                                    ,'' [TB009]
                                    ,[商品編號] [TB010]
                                    ,'******' [TB011]
                                    ,'**********' [TB012]
                                    ,MB001 [TB013]
                                    ,'' [TB014]
                                    ,0 [TB015]
                                    ,[單價] [TB016]
                                    ,0 [TB017]
                                    ,[單價] [TB018]
                                    ,[數量] [TB019]
                                    ,0 [TB020]
                                    ,0 [TB021]
                                    ,0 [TB022]
                                    ,0 [TB023]
                                    ,0 [TB024]
                                    ,0 [TB025]
                                    ,0 [TB026]
                                    ,0 [TB027]
                                    ,0 [TB028]
                                    ,0 [TB029]
                                    ,0 [TB030]
                                    ,([明細金額]-CONVERT(INT,[明細金額]*0.05)) [TB031]
                                    ,CONVERT(INT,[明細金額]*0.05) [TB032]
                                    ,([明細金額]) [TB033]
                                    ,'' [TB034]
                                    ,'' [TB035]
                                    ,'' [TB036]
                                    ,'' [TB037]
                                    ,'' [TB038]
                                    ,'' [TB039]
                                    ,'' [TB040]
                                    ,'' [TB041]
                                    ,'1' [TB042]
                                    ,'N' [TB043]
                                    ,'' [TB044]
                                    ,'' [TB045]
                                    ,'' [TB046]
                                    ,'' [TB047]
                                    ,'1' [TB048]
                                    ,'' [TB049]
                                    ,'' [TB050]
                                    ,'' [TB051]
                                    ,'' [TB052]
                                    ,0 [TB053]
                                    ,0 [TB054]
                                    ,'Y' [TB055]
                                    ,'' [TB056]
                                    ,'' [TB057]
                                    ,'' [TB058]
                                    ,'' [TB059]
                                    ,'' [TB060]
                                    ,'' [TB061]
                                    ,0 [TB062]
                                    ,0 [TB063]
                                    ,0 [TB064]
                                    ,0 [TB065]
                                    ,0 [TB066]
                                    ,0 [TB067]
                                    ,'' [TB068]
                                    ,'' [TB069]
                                    ,'' [TB070]
                                    ,'' [TB071]
                                    ,'' [TB072]
                                    ,'' [TB073]
                                    ,'' [TB074]
                                    ,'' [TB075]
                                    ,0 [TB076]
                                    ,'' [TB077]
                                    ,'' [TB078]
                                    ,'' [TB079]
                                    ,0 [TB080]
                                    ,0 [TB081]
                                    ,'1' [TB082]
                                    ,0 [TB083]
                                    ,'' [TB084]
                                    ,'' [TB085]
                                    ,'' [TB086]
                                    ,'' [TB087]
                                    ,'' [TB088]
                                    ,'' [TB089]
                                    ,'' [TB090]
                                    ,'' [UDF01]
                                    ,'' [UDF02]
                                    ,'' [UDF03]
                                    ,'' [UDF04]
                                    ,'' [UDF05]
                                    ,0 [UDF06]
                                    ,0 [UDF07]
                                    ,0 [UDF08]
                                    ,0 [UDF09]
                                    ,0 [UDF10]
                                    FROM [192.168.1.105].[TKMK].[dbo].[TBJabezPOS],[192.168.1.105].[TK].dbo.INVMB
                                    WHERE 1=1
                                    AND [商品編號]=MB001
                                    AND REPLACE([TBJabezPOS].TA001+[TBJabezPOS].TA002+[TBJabezPOS].TA003+[TBJabezPOS].TA006+[TBJabezPOS].TB007,' ','' )NOT IN (SELECT REPLACE(TB001+TB002+TB003+TB006+TB007, ' ','') FROM [COSMOS_POS].dbo.POSTB  WHERE TB002 IN ('106702')  AND TB003 IN ('900') )

                                    INSERT INTO  [COSMOS_POS].dbo.POSTC
                                    (
                                    [COMPANY]
                                    ,[CREATOR]
                                    ,[USR_GROUP]
                                    ,[CREATE_DATE]
                                    ,[MODIFIER]
                                    ,[MODI_DATE]
                                    ,[FLAG]
                                    ,[CREATE_TIME]
                                    ,[MODI_TIME]
                                    ,[TRANS_TYPE]
                                    ,[TRANS_NAME]
                                    ,[sync_date]
                                    ,[sync_time]
                                    ,[sync_mark]
                                    ,[sync_count]
                                    ,[DataUser]
                                    ,[DataGroup]
                                    ,[TC001]
                                    ,[TC002]
                                    ,[TC003]
                                    ,[TC004]
                                    ,[TC005]
                                    ,[TC006]
                                    ,[TC007]
                                    ,[TC008]
                                    ,[TC009]
                                    ,[TC010]
                                    ,[TC011]
                                    ,[TC012]
                                    ,[TC013]
                                    ,[TC014]
                                    ,[TC015]
                                    ,[TC016]
                                    ,[TC017]
                                    ,[TC018]
                                    ,[TC019]
                                    ,[TC020]
                                    ,[TC021]
                                    ,[TC022]
                                    ,[TC023]
                                    ,[TC024]
                                    ,[TC025]
                                    ,[TC026]
                                    ,[TC027]
                                    ,[TC028]
                                    ,[TC029]
                                    ,[TC030]
                                    ,[TC031]
                                    ,[TC032]
                                    ,[TC033]
                                    ,[TC034]
                                    ,[TC035]
                                    ,[TC036]
                                    ,[TC037]
                                    ,[TC038]
                                    ,[TC039]
                                    ,[TC040]
                                    ,[TC041]
                                    ,[TC042]
                                    ,[TC043]
                                    ,[TC044]
                                    ,[TC045]
                                    ,[TC046]
                                    ,[TC047]
                                    ,[TC048]
                                    ,[TC049]
                                    ,[TC050]
                                    ,[TC051]
                                    ,[TC052]
                                    ,[TC053]
                                    ,[TC054]
                                    ,[TC055]
                                    ,[TC056]
                                    ,[TC057]
                                    ,[UDF01]
                                    ,[UDF02]
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                    ,[UDF06]
                                    ,[UDF07]
                                    ,[UDF08]
                                    ,[UDF09]
                                    ,[UDF10]
                                    ) 
                                    SELECT
                                    'TK' [COMPANY]
                                    ,'150073' [CREATOR]
                                    ,'' [USR_GROUP]
                                    ,[TA001] [CREATE_DATE]
                                    ,'' [MODIFIER]
                                    ,'' [MODI_DATE]
                                    ,0 [FLAG]
                                    ,[時間] [CREATE_TIME]
                                    ,'' [MODI_TIME]
                                    ,'' [TRANS_TYPE]
                                    ,'' [TRANS_NAME]
                                    ,'' [sync_date]
                                    ,'' [sync_time]
                                    ,'N' [sync_mark]
                                    ,0 [sync_count]
                                    ,'' [DataUser]
                                    ,'' [DataGroup]
                                    ,[TA001] [TC001]
                                    ,[TA002] [TC002]
                                    ,[TA003] [TC003]
                                    ,[TA001] [TC004]
                                    ,[時間] [TC005]
                                    ,[TA006] [TC006]
                                    ,[TC007] [TC007]
                                    ,'0001' [TC008]
                                    ,SUM([明細金額])[TC009]
                                    ,'N' [TC010]
                                    ,'' [TC011]
                                    ,'6' [TC012]
                                    ,'' [TC013]
                                    ,'' [TC014]
                                    ,'' [TC015]
                                    ,'1' [TC016]
                                    ,'' [TC017]
                                    ,'Y' [TC018]
                                    ,'' [TC019]
                                    ,'' [TC020]
                                    ,'' [TC021]
                                    ,'' [TC022]
                                    ,0 [TC023]
                                    ,0 [TC024]
                                    ,'' [TC025]
                                    ,'' [TC026]
                                    ,'' [TC027]
                                    ,'' [TC028]
                                    ,'' [TC029]
                                    ,'' [TC030]
                                    ,0 [TC031]
                                    ,0 [TC032]
                                    ,0 [TC033]
                                    ,0 [TC034]
                                    ,0 [TC035]
                                    ,0 [TC036]
                                    ,0 [TC037]
                                    ,'' [TC038]
                                    ,0 [TC039]
                                    ,0 [TC040]
                                    ,0 [TC041]
                                    ,0 [TC042]
                                    ,0 [TC043]
                                    ,'' [TC044]
                                    ,'' [TC045]
                                    ,'' [TC046]
                                    ,'' [TC047]
                                    ,'' [TC048]
                                    ,'' [TC049]
                                    ,'' [TC050]
                                    ,'' [TC051]
                                    ,'' [TC052]
                                    ,'' [TC053]
                                    ,'' [TC054]
                                    ,'' [TC055]
                                    ,'N' [TC056]
                                    ,'N' [TC057]
                                    ,'' [UDF01]
                                    ,'' [UDF02]
                                    ,'' [UDF03]
                                    ,'' [UDF04]
                                    ,'' [UDF05]
                                    ,0 [UDF06]
                                    ,0 [UDF07]
                                    ,0 [UDF08]
                                    ,0 [UDF09]
                                    ,0 [UDF10]
                                    FROM [192.168.1.105].[TKMK].[dbo].[TBJabezPOS]
                                    WHERE REPLACE([TBJabezPOS].TA001+[TBJabezPOS].TA002+[TBJabezPOS].TA003+[TBJabezPOS].TA006+[TBJabezPOS].TC007,' ','') NOT IN (SELECT REPLACE(TC001+TC002+TC003+TC006+TC007,' ','') FROM [COSMOS_POS].dbo.POSTC  WHERE TC002 IN ('106702')  AND TC003 IN ('900') )
                                    GROUP BY [TA001],[TA002],[TA003],[TA006],[時間],[發票],[TC007]


                                    ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 200;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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

        public void DOWNLOAD_EXCEL()
        {
            string url = @"\\192.168.1.109\prog更新\TKMK\REPORT\rcPOS9002_20230515v2.xls";            

            using (FolderBrowserDialog FDB = new FolderBrowserDialog())
            {
                DialogResult RESULT = FDB.ShowDialog();
                if(RESULT==DialogResult.OK&&!string.IsNullOrWhiteSpace(FDB.SelectedPath))
                {
                    string PATH = FDB.SelectedPath;

                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(url, PATH+"\\"+"text.xls");

                        MessageBox.Show("完成");
                    }
                }
            }
            
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMM"), comboBox1.Text);

        }
        private void button4_Click(object sender, EventArgs e)
        {
            CHECKADDDATA();

            Search(dateTimePicker1.Value.ToString("yyyyMM"), comboBox1.Text);
            MessageBox.Show("完成");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //更新TA001、TA002、TA003
            UPDATE_TBJabezPOS_TA001TA002TA003(textBox1.Text);
            //更新TA006
            UPDATE_TBJabezPOS_TA006();
            //更新TB007、TC007
            UPDATE_TBJabezPOS_TB007TC007();


            Search(dateTimePicker1.Value.ToString("yyyy/MM"), comboBox1.Text);
            MessageBox.Show("完成");


        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADD_ERP_POSTAPOSTBPOSTC(); 

            Search(dateTimePicker1.Value.ToString("yyyy/MM"), comboBox1.Text);
            MessageBox.Show("完成");

        }

        private void button5_Click(object sender, EventArgs e)
        {
            DOWNLOAD_EXCEL();
        }

        #endregion


    }
}
