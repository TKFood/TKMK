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
        public void CHECKADDDATA()
        {
            //IEnumerable<DataRow> tempExcept = null;

            //DataTable DT1 = SEARCHTBJabezPOS();
            DataTable DT2 = IMPORTEXCEL();

            //找DataTable差集
            //要有相同的欄位名稱
            //找DataTable差集
            //如果兩個datatable中有部分欄位相同，可以使用Contains比較　　
            //var tempExcept = from r in DT2.AsEnumerable()
            //                 where
            //                 !(from rr in DT1.AsEnumerable() select rr.Field<string>("訂單編號")).Contains(
            //                 r.Field<string>("訂單編號"))
            //                 select r;


            //var tempExcept = DT2.AsEnumerable();

            //if (tempExcept.Count() > 0)
            //{
            //    //差集集合
            //    DataTable dt3 = tempExcept.CopyToDataTable();

            //    INSERTINTOTEMP91APPCOP(dt3);
            //}
        }

        public DataTable IMPORTEXCEL()
        {
            //記錄選到的檔案路徑
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;*.csv;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {
                return null;
            }
            if (dr == DialogResult.Cancel)
            {
                return null;
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
                    _path = @"F:\銷售明細.csv";
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

                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(dtExcelData);
                DataTable Exceldt = dtExcelData;

                //把第一列的欄位名移除
                //Exceldt.Rows[0].Delete();

                if (Exceldt.Rows.Count > 0)
                {
                    return Exceldt;
                }
                else
                {
                    return null;
                }


            }
            catch (Exception ex)
            {
                return null;
                //MessageBox.Show(string.Format("錯誤:{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        #endregion

        #region BUTTON

        private void button4_Click(object sender, EventArgs e)
        {
            CHECKADDDATA();

            MessageBox.Show("完成");
        }
        #endregion
    }
}
