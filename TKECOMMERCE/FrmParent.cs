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

namespace TKECOMMERCE
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
    }
}
