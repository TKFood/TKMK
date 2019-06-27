using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Reflection;
using System.Diagnostics;

namespace TKMK
{
    public partial class FrmParent : Form
    {
        SqlConnection conn;
        MenuStrip MnuStrip;
        ToolStripMenuItem MnuStripItem;
        string UserName;

        public FrmParent()
        {
            InitializeComponent();
        }

        public FrmParent(string txt_UserName)
        {
            InitializeComponent();
            UserName = txt_UserName;
        }

        private void FrmParent_Load(object sender, EventArgs e)
        {
            // To make this Form the Parent Form
            this.IsMdiContainer = true;

            //Creating object of MenuStrip class
            MnuStrip = new MenuStrip();

            //Placing the control to the Form
            this.Controls.Add(MnuStrip);

            String connectionString;
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            conn = new SqlConnection(connectionString);
            String Sequel = "SELECT MAINMNU,MENUPARVAL,STATUS FROM MNU_PARENT";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, conn);
            DataTable dt = new DataTable();
            conn.Open();
            da.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                MnuStripItem = new ToolStripMenuItem(dr["MAINMNU"].ToString());
                SubMenu(MnuStripItem, dr["MENUPARVAL"].ToString());
                MnuStrip.Items.Add(MnuStripItem);
            }
            // The Form.MainMenuStrip property determines the merge target.
            this.MainMenuStrip = MnuStrip;
        }

        public void SubMenu(ToolStripMenuItem mnu, string submenu)
        {
            StringBuilder Seqchild = new StringBuilder();
            Seqchild.AppendFormat("SELECT FRM_NAME FROM MNU_SUBMENU ,MNU_SUBMENULogin WHERE MNU_SUBMENU.FRM_CODE=MNU_SUBMENULogin.FRM_CODE AND  MNU_SUBMENULogin.UserName='{0}' AND MENUPARVAL='{1}'", UserName.ToString(), submenu.ToString());
            //Seqchild.AppendFormat( "SELECT FRM_NAME FROM MNU_SUBMENU ,MNU_SUBMENULogin WHERE MNU_SUBMENU.FRM_CODE=MNU_SUBMENULogin.FRM_CODE AND  MNU_SUBMENULogin.UserName='1' AND MENUPARVAL='1'");
            SqlDataAdapter dachildmnu = new SqlDataAdapter(Seqchild.ToString(), conn);
            DataTable dtchild = new DataTable();
            dachildmnu.Fill(dtchild);

            foreach (DataRow dr in dtchild.Rows)
            {
                ToolStripMenuItem SSMenu = new ToolStripMenuItem(dr["FRM_NAME"].ToString(), null, new EventHandler(ChildClick));
                mnu.DropDownItems.Add(SSMenu);
            }
        }

        private void ChildClick(object sender, EventArgs e)
        {

        }

        private void FrmParent_FormClosed(object sender, FormClosedEventArgs e)
        {
            //=====偵測執行中的外部程式並關閉=====
            Process[] MyProcess = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (MyProcess.Length > 0)
                MyProcess[0].Kill(); //關閉執行中的程式


        }
    }

}
