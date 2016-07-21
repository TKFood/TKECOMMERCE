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

namespace TKECOMMERCE
{
    public partial class FrmMPRECOPTC : Form
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
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;
        string NowDB = "test";

        public FrmMPRECOPTC()
        {
            InitializeComponent();
            SetMyCustomFormat();
        }

        #region FUNCTION
        public void SetMyCustomFormat()
        {
            // Set the Format type and the CustomFormat string.
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy - MM";
        }

        public void Search()
        {
            try
            {

                if (!string.IsNullOrEmpty(dateTimePicker1.Text.ToString()) )
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSqlQuery.AppendFormat("{0}",dateTimePicker1.Value.ToString("yyyyMM"));
                    sbSql.AppendFormat(@"SELECT  YEARMONTH AS '年月', ZTKECOMMERCEFrmMPRECOPTC.MB001 AS '品號', ZTKECOMMERCEFrmMPRECOPTC.MB002 AS '品名', PREOrderNum AS '數量',MB004 AS '單位' FROM  [{0}].[dbo].ZTKECOMMERCEFrmMPRECOPTC  WITH (NOLOCK) ,[{1}].[dbo].INVMB  WITH (NOLOCK) WHERE ZTKECOMMERCEFrmMPRECOPTC.MB001=INVMB.MB001 AND YEARMONTH='{2}'  ", NowDB, NowDB, sbSqlQuery.ToString());

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "TEMPds");
                    sqlConn.Close();


                    if (ds.Tables["TEMPds"].Rows.Count == 0)
                    {
                        
                    }
                    else
                    {                       

                        dataGridView1.DataSource = ds.Tables["TEMPds"];
                        dataGridView1.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public void OpenFile()
        {            
            file.ShowDialog();
            file.Filter =
                "請開啟EXCEL檔 files (*XlSX" +
                 "All Files (*.*)|*.*";
            file.Title = "請開啟EXCEL檔";

            strFilePath = file.FileName.ToString();
            textBox1.Text = file.SafeFileName;      

        }

        public void ExcelImport()
        {
            try
            {
                IWorkbook workbook;
                using (FileStream stream = new FileStream(strFilePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(stream);
                }

                ISheet sheet = workbook.GetSheetAt(0); // zero-based index of your target sheet
                dt = new DataTable(sheet.SheetName);

                // write header row
                IRow headerRow = sheet.GetRow(0);
                foreach (ICell headerCell in headerRow)
                {
                    dt.Columns.Add(headerCell.ToString());
                }

                // write the rest
                int rowIndex = 0;
                foreach (IRow row in sheet)
                {
                    // skip header row
                    if (rowIndex++ == 0) continue;
                    DataRow dataRow = dt.NewRow();
                    dataRow.ItemArray = row.Cells.Select(c => c.ToString()).ToArray();
                    dt.Rows.Add(dataRow);
                }

          
                //add ZTKECOMMERCEFrmMPRECOPTC
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" ");               
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sbSql.AppendFormat(" INSERT INTO [{0}].[dbo].[ZTKECOMMERCEFrmMPRECOPTC] ", NowDB);
                    sbSql.Append(" (YEARMONTH, MB001,MB002,  PREOrderNum )");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','NA','{2}')", dt.Rows[i]["年月"].ToString(), dt.Rows[i]["品號"].ToString(), dt.Rows[i]["數量"].ToString());
                    sbSql.Append(" ");
                }
                sbSql.AppendFormat(" UPDATE [{0}].[dbo].[ZTKECOMMERCEFrmMPRECOPTC] SET ZTKECOMMERCEFrmMPRECOPTC.MB002=INVMB.MB002 FROM [{1}].[dbo].[INVMB] WHERE ZTKECOMMERCEFrmMPRECOPTC.MB001=INVMB.MB001 AND ZTKECOMMERCEFrmMPRECOPTC.MB002='NA'", NowDB, NowDB);
                sbSql.Append(" ");
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
            //dataGridView1.DataSource = dt;
            Search();
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelImport();
        }
        #endregion


    }
}
