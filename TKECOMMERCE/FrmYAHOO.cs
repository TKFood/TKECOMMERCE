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
using TKITDLL;

namespace TKECOMMERCE
{
    public partial class FrmYAHOO : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        DataTable dt = new DataTable();

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        private Report report1;

        public FrmYAHOO()
        {
            InitializeComponent();
        }


        #region FUNCTION
        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
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
            if (dataGridView1.Rows.Count>=1)
            {
                Bulk_Insert();

                //foreach (DataGridViewRow row in dataGridView1.Rows)
                //{
                //    using (SqlConnection con = new SqlConnection(connectionString))
                //    {
                //        using (SqlCommand cmd = new SqlCommand("INSERT INTO [TKECOMMERCE].[dbo].[YAHOO] (ID) VALUES(@ID)", con))
                //        {
                //            cmd.Parameters.AddWithValue("@ID", row.Cells["訂單編號"].Value);
                //            con.Open();
                //            cmd.ExecuteNonQuery();
                //            con.Close();
                //        }
                //    }

                //    //MessageBox.Show(row.Cells["訂單編號"].Value.ToString());
                //    //More code here
                //}
            }
        }

        protected void Bulk_Insert()
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[30] {
                new DataColumn("ID", typeof(string)),
                new DataColumn("NAME", typeof(string)),
                new DataColumn("SERNO", typeof(string)),
                new DataColumn("PAYKIND", typeof(string)),
                new DataColumn("RECIVER", typeof(string)),
                new DataColumn("POST", typeof(string)),
                new DataColumn("ADDER", typeof(string)),
                new DataColumn("ADDDATE", typeof(DateTime)),
                new DataColumn("SHIPDATE", typeof(DateTime)),
                new DataColumn("SHOPDATE", typeof(DateTime)),
                new DataColumn("KIND", typeof(string)),
                new DataColumn("DELIVERY", typeof(string)),
                new DataColumn("YNO", typeof(string)),
                new DataColumn("PNO", typeof(string)),
                new DataColumn("PNAME", typeof(string)),
                new DataColumn("REMARK", typeof(string)),
                new DataColumn("SPEC", typeof(string)),
                new DataColumn("QUANTITY", typeof(string)),
                new DataColumn("TMONEY", typeof(string)),
                new DataColumn("STATES", typeof(string)),
                new DataColumn("INMONEYDATE", typeof(DateTime)),
                new DataColumn("TELDAY", typeof(string)),
                new DataColumn("TELNIGHT", typeof(string)),
                new DataColumn("MOBILE", typeof(string)),
                new DataColumn("TAX", typeof(string)),
                new DataColumn("BONUS", typeof(string)),
                new DataColumn("BONUSMONEY", typeof(string)),
                new DataColumn("DISCOUNTACT", typeof(string)),
                new DataColumn("DISCOUNTCODE", typeof(string)),
                new DataColumn("DISCOUNTMONEY", typeof(string))
            });

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string ID = row.Cells["訂單編號"].Value.ToString();
                    string NAME = row.Cells["訂購人"].Value.ToString();
                    string SERNO = row.Cells["交易序號"].Value.ToString();
                    string PAYKIND = row.Cells["付款別"].Value.ToString();
                    string RECIVER = row.Cells["收件人姓名"].Value.ToString();
                    string POST = row.Cells["收件人郵遞區號"].Value.ToString();
                    string ADDER = row.Cells["收件人地址"].Value.ToString();
                    DateTime ADDDATE = Convert.ToDateTime(row.Cells["轉單日期"].Value.ToString());
                    DateTime SHIPDATE = Convert.ToDateTime(row.Cells["最晚出貨日"].Value.ToString());
                    DateTime SHOPDATE = new DateTime(1911, 1, 1);
                    string KIND = row.Cells["商品類型"].Value.ToString();
                    string DELIVERY = row.Cells["物流設定"].Value.ToString();
                    string YNO = row.Cells["商品編號"].Value.ToString();
                    string PNO = row.Cells["店家商品料號"].Value.ToString();
                    string PNAME = row.Cells["商品名稱"].Value.ToString();
                    string REMARK = row.Cells["購物車備註"].Value.ToString();
                    string SPEC = row.Cells["商品規格"].Value.ToString();
                    string QUANTITY = row.Cells["數量"].Value.ToString();
                    string TMONEY = row.Cells["金額小計"].Value.ToString();
                    string STATES = row.Cells["訂單狀態"].Value.ToString();
                    DateTime INMONEYDATE = new DateTime(1911, 1, 1);
                    string TELDAY = row.Cells["收件人電話(日)"].Value.ToString();
                    string TELNIGHT = row.Cells["收件人電話(夜)"].Value.ToString();
                    string MOBILE = row.Cells["收件人行動電話"].Value.ToString();
                    string TAX = row.Cells["商品稅別"].Value.ToString();
                    string BONUS = row.Cells["超贈點點數"].Value.ToString();
                    string BONUSMONEY = row.Cells["超贈點折抵金額"].Value.ToString();
                    string DISCOUNTACT = row.Cells["折扣碼活動編號"].Value.ToString();
                    string DISCOUNTCODE = row.Cells["折扣碼"].Value.ToString();
                    string DISCOUNTMONEY = row.Cells["折扣碼折抵金額"].Value.ToString();

                    if (string.IsNullOrEmpty(row.Cells["店家出貨日"].Value.ToString()))
                    {
                        SHOPDATE = new DateTime(1911, 1, 1);
                    }
                    else
                    {
                        SHOPDATE = Convert.ToDateTime(row.Cells["店家出貨日"].Value.ToString());
                    }

                    if (string.IsNullOrEmpty(row.Cells["入帳日"].Value.ToString()))
                    {
                        INMONEYDATE = new DateTime(1911, 1, 1);
                    }
                    else
                    {
                        INMONEYDATE = Convert.ToDateTime(row.Cells["入帳日"].Value.ToString());
                    }

                    dt.Rows.Add(ID, NAME, SERNO, PAYKIND, RECIVER, POST, ADDER,ADDDATE,SHIPDATE,SHOPDATE, KIND, DELIVERY,YNO, PNO, PNAME,REMARK, SPEC, QUANTITY, TMONEY,STATES,INMONEYDATE,TELDAY,TELNIGHT ,MOBILE, TAX ,BONUS,BONUSMONEY, DISCOUNTACT, DISCOUNTCODE, DISCOUNTMONEY);
                }

                if (dt.Rows.Count > 0)
                {

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = new SqlConnection(sqlsb.ConnectionString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name
                            sqlBulkCopy.DestinationTableName = "[TKECOMMERCE].[dbo].[YAHOO]";

                            //[OPTIONAL]: Map the DataTable columns with that of the database table
                            sqlBulkCopy.ColumnMappings.Add("ID", "ID");
                            sqlBulkCopy.ColumnMappings.Add("NAME", "NAME");
                            sqlBulkCopy.ColumnMappings.Add("SERNO", "SERNO");
                            sqlBulkCopy.ColumnMappings.Add("PAYKIND", "PAYKIND");
                            sqlBulkCopy.ColumnMappings.Add("RECIVER", "RECIVER");
                            sqlBulkCopy.ColumnMappings.Add("POST", "POST");
                            sqlBulkCopy.ColumnMappings.Add("ADDER", "ADDER");
                            sqlBulkCopy.ColumnMappings.Add("ADDDATE", "ADDDATE");
                            sqlBulkCopy.ColumnMappings.Add("SHIPDATE", "SHIPDATE");
                            sqlBulkCopy.ColumnMappings.Add("SHOPDATE", "SHOPDATE");
                            sqlBulkCopy.ColumnMappings.Add("KIND", "KIND");
                            sqlBulkCopy.ColumnMappings.Add("DELIVERY", "DELIVERY");
                            sqlBulkCopy.ColumnMappings.Add("YNO", "YNO");
                            sqlBulkCopy.ColumnMappings.Add("PNO", "PNO");
                            sqlBulkCopy.ColumnMappings.Add("PNAME", "PNAME");
                            sqlBulkCopy.ColumnMappings.Add("REMARK", "REMARK");
                            sqlBulkCopy.ColumnMappings.Add("SPEC", "SPEC");
                            sqlBulkCopy.ColumnMappings.Add("QUANTITY", "QUANTITY");
                            sqlBulkCopy.ColumnMappings.Add("TMONEY", "TMONEY");
                            sqlBulkCopy.ColumnMappings.Add("STATES", "STATES");
                            sqlBulkCopy.ColumnMappings.Add("INMONEYDATE", "INMONEYDATE");
                            sqlBulkCopy.ColumnMappings.Add("TELDAY", "TELDAY");
                            sqlBulkCopy.ColumnMappings.Add("TELNIGHT", "TELNIGHT");
                            sqlBulkCopy.ColumnMappings.Add("MOBILE", "MOBILE");
                            sqlBulkCopy.ColumnMappings.Add("TAX", "TAX");
                            sqlBulkCopy.ColumnMappings.Add("BONUS", "BONUS");
                            sqlBulkCopy.ColumnMappings.Add("BONUSMONEY", "BONUSMONEY");
                            sqlBulkCopy.ColumnMappings.Add("DISCOUNTACT", "DISCOUNTACT");
                            sqlBulkCopy.ColumnMappings.Add("DISCOUNTCODE", "DISCOUNTCODE");
                            sqlBulkCopy.ColumnMappings.Add("DISCOUNTMONEY", "DISCOUNTMONEY");


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
        public void SETFASTREPORT()
        {
            report1 = new Report();
            report1.Load(@"REPORT\YAHOO訂單.frx");
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;            

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            DateTime dt = Convert.ToDateTime(dateTimePicker2.Value);
            dt = dt.AddDays(1);

            report1.SetParameterValue("P2", dt.ToString("yyyyMMdd"));        

            report1.Preview = previewControl1;
            report1.Show();
        }

        public void SETNULLDT()
        {
            //dt = null;
            dataGridView1.DataSource = null;
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
