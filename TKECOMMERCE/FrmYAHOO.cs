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
    public partial class FrmYAHOO : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        DataTable dt = new DataTable();

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;

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
                dt.Columns.AddRange(new DataColumn[2] {
                new DataColumn("ID", typeof(string)),
                new DataColumn("NAME", typeof(string))
            });

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string ID = row.Cells["訂單編號"].Value.ToString();
                    string NAME = row.Cells["訂購人"].Value.ToString();

                    dt.Rows.Add(ID, NAME);
                }

                if (dt.Rows.Count > 0)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name
                            sqlBulkCopy.DestinationTableName = "[TKECOMMERCE].[dbo].[YAHOO]";

                            //[OPTIONAL]: Map the DataTable columns with that of the database table
                            sqlBulkCopy.ColumnMappings.Add("ID", "ID");
                            sqlBulkCopy.ColumnMappings.Add("NAME", "NAME");

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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ImportDB();
        }

        #endregion


    }
}
