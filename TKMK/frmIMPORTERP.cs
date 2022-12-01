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

        string _path = null;
        DataTable EXCEL = null;


        public frmIMPORTERP()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void Search(string YYYYMM)
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
                                    ,[TA001]
                                    ,[TA002]
                                    ,[TA003]
                                    ,[TA006]
                                    ,[TB007]
                                    ,[TC007]
                                     ,RIGHT('00000'+CAST(row_number() OVER(PARTITION BY [TA001],[TA002],[TA003] ORDER BY [TA001],[TA002],[TA003]) AS nvarchar(10)),5)  AS SEQ
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE [日期] LIKE '%{0}%'
                                    ORDER BY [機台],[日期],[序號],[時間]

                                         ", YYYYMM);

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
             IMPORTEXCEL();

            //檢查暫存中的新資料
            DataTable NEWTDATATABLE = SEARCHNEWDATA();
            //匯入到TBJabezPOS中
            if(NEWTDATATABLE.Rows.Count>0)
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
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;*.csv;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {

            }
            if (dr == DialogResult.Cancel)
            {
                
            }


            textBox3.Text = od.FileName.ToString();
            _path = od.FileName.ToString();

            try
            {
                //  ExcelConn(_path);
                //找出不同excel的格式，設定連接字串
                //xls跟非xls
                string constr = null;
                string CHECKEXCELFORMAT = _path.Substring(_path.Length - 4, 4);


                if (CHECKEXCELFORMAT.CompareTo(".xls") == 0)
                {
                    constr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";  
                }
                else if (CHECKEXCELFORMAT.CompareTo(".csv") == 0)
                {
                    //_path = @"F:\銷售明細.csv";
                    //constr = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = F:\銷售明細.csv; Extended Properties = 'TEXT;IMEX=1;HDR=Yes;FMT=Delimited;CharacterSet=UNICODE;'";
                    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path.Remove(_path.LastIndexOf("\\") + 1) + ";Extended Properties='Text;FMT=Delimited;HDR=YES;'";
                }
                else
                {
                    constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _path + ";Extended Properties='Excel 12.0;HDR=NO';";  
                }

                //找出excel的第1張分頁名稱，用query中                
                OleDbConnection Econ = new OleDbConnection(constr);
                Econ.Open();



                DataTable excelShema = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string firstSheetName = null;
                //找到csv的正確table表
                if (CHECKEXCELFORMAT.CompareTo(".csv") == 0)
                {
                    for (int i = 0; i < excelShema.Rows.Count; i++)
                    {
                        firstSheetName = Convert.ToString(excelShema.Rows[i]["TABLE_NAME"]);

                        if (firstSheetName.Contains("csv"))
                        {
                            break;
                        }
                    }
                }
                else
                {
                     excelShema.Rows[0]["TABLE_NAME"].ToString();
                }
                   
                   

                string Query = string.Format("Select * FROM [{0}]", firstSheetName);
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);


                DataTable dtExcelData = new DataTable();
                dtExcelData.Columns.Add("機台", typeof(string));
                dtExcelData.Columns.Add("日期", typeof(string));
                dtExcelData.Columns.Add("序號", typeof(string));
                dtExcelData.Columns.Add("時間", typeof(string));
                dtExcelData.Columns.Add("商品編號", typeof(string));
                dtExcelData.Columns.Add("商品規格", typeof(string));
                dtExcelData.Columns.Add("單價", typeof(decimal));
                dtExcelData.Columns.Add("數量", typeof(decimal));
                dtExcelData.Columns.Add("小計", typeof(decimal));

                DataTable Exceldt = new DataTable();
                Exceldt.Columns.Add("機台", typeof(string));
                Exceldt.Columns.Add("日期", typeof(string));
                Exceldt.Columns.Add("序號", typeof(string));
                Exceldt.Columns.Add("時間", typeof(string));
                Exceldt.Columns.Add("商品編號", typeof(string));
                Exceldt.Columns.Add("商品規格", typeof(string));
                Exceldt.Columns.Add("單價", typeof(decimal));
                Exceldt.Columns.Add("數量", typeof(decimal));
                Exceldt.Columns.Add("小計", typeof(decimal));

                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(dtExcelData);

                //轉換日期為文字
                foreach(DataRow  DR in dtExcelData.Rows)
                {
                    DataRow NEWDR = Exceldt.NewRow();
                    NEWDR["機台"] = DR["機台"].ToString();
                    NEWDR["日期"]= Convert.ToDateTime(DR["日期"].ToString()).ToString("yyyy/MM/dd");
                    NEWDR["序號"] = DR["序號"].ToString();
                    NEWDR["時間"] = Convert.ToDateTime(DR["時間"].ToString()).ToString("HH:mm:ss");
                    NEWDR["商品編號"] = DR["商品編號"].ToString();
                    NEWDR["商品規格"] = DR["商品規格"].ToString();
                    NEWDR["單價"] = DR["單價"].ToString();
                    NEWDR["數量"] = DR["數量"].ToString();
                    NEWDR["小計"] = DR["小計"].ToString();

                    Exceldt.Rows.Add(NEWDR);
                }

                //DataTable Exceldt = dtExcelData;

                //把第一列的欄位名移除
                //Exceldt.Rows[0].Delete();

                if (Exceldt.Rows.Count > 0)
                {
                    ADD_TO_TBJabezPOS_TEMP(Exceldt);
                }
                else
                {
                    
                }


            }
            catch (Exception ex)
            {
                
                //MessageBox.Show(string.Format("錯誤:{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
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
                bulkCopy.ColumnMappings.Add("機台", "機台");
                bulkCopy.ColumnMappings.Add("日期", "日期");
                bulkCopy.ColumnMappings.Add("序號", "序號");
                bulkCopy.ColumnMappings.Add("時間", "時間");
                bulkCopy.ColumnMappings.Add("商品編號", "商品編號");
                bulkCopy.ColumnMappings.Add("商品規格", "商品規格");
                bulkCopy.ColumnMappings.Add("單價", "單價");
                bulkCopy.ColumnMappings.Add("數量", "數量");
                bulkCopy.ColumnMappings.Add("小計", "小計");
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
                                     [機台]
                                    ,[日期]
                                    ,[序號]
                                    ,[時間]
                                    ,[商品編號]
                                    ,[商品規格]
                                    ,[單價]
                                    ,[數量]
                                    ,[小計]
                                    FROM [TKMK].[dbo].[TBJabezPOS_TEMP]
                                    WHERE [機台]+[日期]+[序號] NOT IN (SELECT [機台]+[日期]+[序號] FROM [TKMK].[dbo].[TBJabezPOS])
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
                bulkCopy.ColumnMappings.Add("機台", "機台");
                bulkCopy.ColumnMappings.Add("日期", "日期");
                bulkCopy.ColumnMappings.Add("序號", "序號");
                bulkCopy.ColumnMappings.Add("時間", "時間");
                bulkCopy.ColumnMappings.Add("商品編號", "商品編號");
                bulkCopy.ColumnMappings.Add("商品規格", "商品規格");
                bulkCopy.ColumnMappings.Add("單價", "單價");
                bulkCopy.ColumnMappings.Add("數量", "數量");
                bulkCopy.ColumnMappings.Add("小計", "小計");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyy/MM"));

        }
        private void button4_Click(object sender, EventArgs e)
        {
            CHECKADDDATA();

            MessageBox.Show("完成");
        }
        #endregion

       
    }
}
