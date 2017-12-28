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
using System.Data.OleDb;

namespace TKECOMMERCE
{
    public partial class FrmMOMO : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        DataTable dt = new DataTable();

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        private Report report1;

        public FrmMOMO()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            string header = "YES";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        dt.Clear();

                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();

                        //Populate DataGridView.
                        dataGridView1.DataSource = dt;
                    }
                }
            }
        }

        public void ImportDB()
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                Bulk_Insert();

            }
        }

        protected void Bulk_Insert()
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[24] {
                new DataColumn("ORDERNO", typeof(string)),
                new DataColumn("DELIVERY", typeof(string)),
                new DataColumn("DMESS", typeof(string)),
                new DataColumn("DCOM", typeof(string)),
                new DataColumn("DNO", typeof(string)),
                new DataColumn("DREQUEST", typeof(string)),
                new DataColumn("PAYDATE", typeof(DateTime)),
                new DataColumn("LASTDATE", typeof(DateTime)),
                new DataColumn("RECIVER", typeof(string)),
                new DataColumn("RECIVERTEL", typeof(string)),
                new DataColumn("RECIVERMOBILE", typeof(string)),
                new DataColumn("ADDER", typeof(string)),
                new DataColumn("MNO", typeof(string)),
                new DataColumn("PNO", typeof(string)),
                new DataColumn("PNAME", typeof(string)),
                new DataColumn("SPEC", typeof(string)),
                new DataColumn("QUANTITY", typeof(int)),
                new DataColumn("TMONEY", typeof(decimal)),
                new DataColumn("PAYKIND", typeof(string)),
                new DataColumn("PAYNUM", typeof(string)),
                new DataColumn("STATES", typeof(string)),
                new DataColumn("ORDERNAME", typeof(string)),
                new DataColumn("ORDERTEL", typeof(string)),
                new DataColumn("ORDERMOBILE", typeof(string))
               
            });

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string ORDERNO = row.Cells["訂單編號"].Value.ToString();
                    string DELIVERY = row.Cells["配送狀態"].Value.ToString();
                    string DMESS = row.Cells["配送訊息"].Value.ToString();
                    string DCOM = row.Cells["物流公司"].Value.ToString();
                    string DNO = row.Cells["配送單號"].Value.ToString();
                    string DREQUEST = row.Cells["客戶配送需求"].Value.ToString();
                    DateTime PAYDATE = Convert.ToDateTime(row.Cells["付款日"].Value.ToString());
                    DateTime LASTDATE = Convert.ToDateTime(row.Cells["最晚出貨日"].Value.ToString());
                    string RECIVER = row.Cells["收件人姓名"].Value.ToString();
                    string RECIVERTEL = row.Cells["電話"].Value.ToString();
                    string RECIVERMOBILE = row.Cells["行動電話"].Value.ToString();
                    string ADDER = row.Cells["地址"].Value.ToString();
                    string MNO = row.Cells["商店品號"].Value.ToString();
                    string PNO = row.Cells["商品編號"].Value.ToString();
                    string PNAME = row.Cells["商品名稱"].Value.ToString();
                    string SPEC = row.Cells["單品規格"].Value.ToString();
                    int QUANTITY = Convert.ToInt32(row.Cells["數量"].Value.ToString());
                    decimal TMONEY = Convert.ToDecimal(row.Cells["成交價"].Value.ToString());
                    string PAYKIND = row.Cells["付款方式"].Value.ToString();
                    string PAYNUM = row.Cells["分期"].Value.ToString();
                    string STATES = row.Cells["商品屬性"].Value.ToString();
                    string ORDERNAME = row.Cells["訂購人姓名"].Value.ToString();
                    string ORDERTEL = row.Cells["電話"].Value.ToString();
                    string ORDERMOBILE = row.Cells["行動電話"].Value.ToString();

                   

                    if (string.IsNullOrEmpty(row.Cells["付款日"].Value.ToString()))
                    {
                        PAYDATE = new DateTime(1911, 1, 1);
                    }
                    else
                    {
                        PAYDATE = Convert.ToDateTime(row.Cells["付款日"].Value.ToString());
                    }

                    if (string.IsNullOrEmpty(row.Cells["最晚出貨日"].Value.ToString()))
                    {
                        LASTDATE = new DateTime(1911, 1, 1);
                    }
                    else
                    {
                        LASTDATE = Convert.ToDateTime(row.Cells["最晚出貨日"].Value.ToString());
                    }

                    
                    dt.Rows.Add(ORDERNO, DELIVERY, DMESS, DCOM, DNO, DREQUEST, PAYDATE, LASTDATE, RECIVER, RECIVERTEL, RECIVERMOBILE, ADDER, MNO, PNO, PNAME, SPEC, QUANTITY, TMONEY, PAYKIND, PAYNUM, STATES, ORDERNAME, ORDERTEL, ORDERMOBILE);
                }

                if (dt.Rows.Count > 0)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name
                            sqlBulkCopy.DestinationTableName = "[TKECOMMERCE].[dbo].[MOMO]";

                            //[OPTIONAL]: Map the DataTable columns with that of the database table
                            sqlBulkCopy.ColumnMappings.Add("ORDERNO", "ORDERNO");
                            sqlBulkCopy.ColumnMappings.Add("DELIVERY", "DELIVERY");
                            sqlBulkCopy.ColumnMappings.Add("DMESS", "DMESS");
                            sqlBulkCopy.ColumnMappings.Add("DCOM", "DCOM");
                            sqlBulkCopy.ColumnMappings.Add("DNO", "DNO");
                            sqlBulkCopy.ColumnMappings.Add("DREQUEST", "DREQUEST");
                            sqlBulkCopy.ColumnMappings.Add("PAYDATE", "PAYDATE");
                            sqlBulkCopy.ColumnMappings.Add("LASTDATE", "LASTDATE");
                            sqlBulkCopy.ColumnMappings.Add("RECIVER", "RECIVER");
                            sqlBulkCopy.ColumnMappings.Add("RECIVERTEL", "RECIVERTEL");
                            sqlBulkCopy.ColumnMappings.Add("RECIVERMOBILE", "RECIVERMOBILE");
                            sqlBulkCopy.ColumnMappings.Add("ADDER", "ADDER");
                            sqlBulkCopy.ColumnMappings.Add("MNO", "MNO");
                            sqlBulkCopy.ColumnMappings.Add("PNO", "PNO");
                            sqlBulkCopy.ColumnMappings.Add("PNAME", "PNAME");
                            sqlBulkCopy.ColumnMappings.Add("SPEC", "SPEC");
                            sqlBulkCopy.ColumnMappings.Add("QUANTITY", "QUANTITY");
                            sqlBulkCopy.ColumnMappings.Add("TMONEY", "TMONEY");
                            sqlBulkCopy.ColumnMappings.Add("PAYKIND", "PAYKIND");
                            sqlBulkCopy.ColumnMappings.Add("PAYNUM", "PAYNUM");
                            sqlBulkCopy.ColumnMappings.Add("STATES", "STATES");
                            sqlBulkCopy.ColumnMappings.Add("ORDERNAME", "ORDERNAME");
                            sqlBulkCopy.ColumnMappings.Add("ORDERTEL", "ORDERTEL");
                            sqlBulkCopy.ColumnMappings.Add("ORDERMOBILE", "ORDERMOBILE");



                            con.Open();
                            sqlBulkCopy.WriteToServer(dt);
                            con.Close();
                        }
                    }
                }

                MessageBox.Show("完成");
            }

            catch
            {
                MessageBox.Show("錯誤");
            }

            finally
            {

            }

        }
        public void SETNULLDT()
        {
            //dt = null;
            dataGridView1.DataSource = null;
        }
        public void SETFASTREPORT()
        {
            report1 = new Report();
            report1.Load(@"REPORT\MOMO訂單.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            DateTime dt = Convert.ToDateTime(dateTimePicker2.Value);
            dt = dt.AddDays(1);

            report1.SetParameterValue("P2", dt.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETNULLDT();
            openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ImportDB();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion


    }
}
