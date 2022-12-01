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
                                            AND  REPLACE(TA001+TA002+TA003+TA006,' ','') IN (SELECT  REPLACE(TA001+TA002+TA003+TA006,' ','') FROM [TK].dbo.POSTATEMP)
                                            ");
                }
                else
                {
                    sbSqlQUERY.AppendFormat(@" 
                                            AND  REPLACE(TA001+TA002+TA003+TA006,' ','') NOT IN (SELECT  REPLACE(TA001+TA002+TA003+TA006,' ','') FROM [TK].dbo.POSTATEMP)
                                            ");
                }

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
                                     
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE [日期] LIKE '%{0}%'
                                    {1}
                                    ORDER BY [機台],[日期],[序號],[時間]

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

        public void UPDATE_TBJabezPOS_TA001TA002TA003()
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
                                    SET [TA001]=CONVERT(NVARCHAR,CONVERT(DATETIME,[日期]),112)
                                    WHERE ISNULL([TA001],'')=''

                                    UPDATE [TKMK].[dbo].[TBJabezPOS]
                                    SET [TA002]=[DVALUES]
                                    FROM [TKMK].[dbo].[TBZPARAS]
                                    WHERE ISNULL([TA002],'')=''
                                    AND [TBZPARAS].KINDS='TBJabezPOS' AND [TBZPARAS].PARASNAMES='店號'

                                    UPDATE [TKMK].[dbo].[TBJabezPOS]
                                    SET [TA003]=[DVALUES]
                                    FROM [TKMK].[dbo].[TBZPARAS]
                                    WHERE ISNULL([TA003],'')=''
                                    AND [TBZPARAS].KINDS='TBJabezPOS' AND [TBZPARAS].PARASNAMES='機號'
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
                                    ,([機台]+[日期]+[序號]) AS 'KEYS'
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE ISNULL(TA006,'')=''
                                    ORDER BY [機台],[日期],[序號],[時間]

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
                                        WHERE ([機台]+[日期]+[序號])='{0}'

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
                                    ,RIGHT('0000'+CAST(row_number() OVER(PARTITION BY [TA001],[TA002],[TA003],[TA006] ORDER BY [TA001],[TA002],[TA003],[TA006]) AS nvarchar(10)),4)  AS SEQ 
                                    ,([TA001]+[TA002]+[TA003]+[TA006]) AS 'KEYS'
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
                                        SET [TB007]='{1}',[TC007]='{1}'
                                        WHERE ([TA001]+[TA002]+[TA003]+[TA006])='{0}'

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

                                    INSERT INTO  [TK].[dbo].[POSTATEMP]
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
                                    ,'' [TA014]
                                    ,0 [TA015]
                                    ,[數量] [TA016]
                                    ,[小計] [TA017]
                                    ,0 [TA018]
                                    ,[小計] [TA019]
                                    ,0 [TA020]
                                    ,CONVERT(INT,[小計]*0.05) [TA021]
                                    ,0 [TA022]
                                    ,0 [TA023]
                                    ,[小計] [TA024]
                                    ,[小計] [TA025]
                                    ,([小計]-CONVERT(INT,[小計]*0.05)) [TA026]
                                    ,CONVERT(INT,[小計]*0.05) [TA027]
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
                                    ,'' [TA041]
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
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE REPLACE([TBJabezPOS].TA001+[TBJabezPOS].TA002+[TBJabezPOS].TA003+[TBJabezPOS].TA006,' ','') NOT IN (SELECT REPLACE(TA001+TA002+TA003+TA006,' ','') FROM [TK].dbo.POSTATEMP)

                                    INSERT INTO  [TK].[dbo].[POSTBTEMP]
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
                                    ,([小計]-CONVERT(INT,[小計]*0.05)) [TB031]
                                    ,CONVERT(INT,[小計]*0.05) [TB032]
                                    ,([小計]) [TB033]
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
                                    FROM [TKMK].[dbo].[TBJabezPOS],[TK].dbo.INVMB
                                    WHERE 1=1
                                    AND [商品編號]=MB001
                                    AND REPLACE([TBJabezPOS].TA001+[TBJabezPOS].TA002+[TBJabezPOS].TA003+[TBJabezPOS].TA006+[TBJabezPOS].TB007,' ','' )NOT IN (SELECT REPLACE(TB001+TB002+TB003+TB006+TB007, ' ','') FROM [TK].dbo.POSTBTEMP)

                                    INSERT INTO  [TK].[dbo].[POSTCTEMP]
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
                                    ,[小計] [TC009]
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
                                    FROM [TKMK].[dbo].[TBJabezPOS]
                                    WHERE REPLACE([TBJabezPOS].TA001+[TBJabezPOS].TA002+[TBJabezPOS].TA003+[TBJabezPOS].TA006+[TBJabezPOS].TC007,' ','') NOT IN (SELECT REPLACE(TC001+TC002+TC003+TC006+TC007,' ','') FROM [TK].dbo.POSTCTEMP)

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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyy/MM"), comboBox1.Text);

        }
        private void button4_Click(object sender, EventArgs e)
        {
            CHECKADDDATA();

            MessageBox.Show("完成");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //更新TA001、TA002、TA003
            UPDATE_TBJabezPOS_TA001TA002TA003();
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
        #endregion


    }
}
