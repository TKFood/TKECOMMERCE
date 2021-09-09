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
using TKITDLL;

namespace TKECOMMERCE
{
    public partial class FrmREDSCCOPTGPCT : Form
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
        string NowDB;
        string NowTable = null;
        int result;

        public FrmREDSCCOPTGPCT()
        {
            InitializeComponent();
            dateTimePicker1.CustomFormat = "yyyyMM";
        }

        #region FUNCTION
        public void ExportExcel(DataSet dsExcel, string Tabelname)
        {
            String NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            dt = dsExcel.Tables[Tabelname];

            ////建立Excel 2007檔案
            //IWorkbook wb = new XSSFWorkbook();
            //ISheet ws;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\銷售完成率{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }
        }

        public void Search()
        {
            NowDB = "TK";

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

            sbSqlQuery.AppendFormat("  [YEARMONTH]='{0}' ", dateTimePicker1.Value.ToString("yyyyMM"));

            if (comboBox1.Text.ToString().Equals("銷售完成率"))
            {
                sbSql.AppendFormat(" SELECT [YEARMONTH] AS '年月',[MB001] AS '品號' ,[MB002] AS '品名',[PREOrderNum] AS '預估量',CAST((SELECT ISNULL(SUM(TH008+TH024),0) FROM [{0}].dbo.COPTH WITH (NOLOCK) WHERE TH004=MB001 AND TH001='A233' AND SUBSTRING(TH002,1,6)=YEARMONTH) AS INT) AS '出貨量',CAST((SELECT ISNULL(SUM(TJ007),0) FROM [{0}].dbo.COPTJ WITH (NOLOCK) WHERE TJ004=MB001 AND TJ001='A246' AND SUBSTRING(TJ002,1,6)=YEARMONTH) AS INT) AS '退貨量',(CAST((SELECT ISNULL(SUM(TH008+TH024),0) FROM [{0}].dbo.COPTH WITH (NOLOCK) WHERE TH004=MB001 AND TH001='A233' AND SUBSTRING(TH002,1,6)=YEARMONTH) AS INT)-CAST((SELECT ISNULL(SUM(TJ007),0) FROM [{0}].dbo.COPTJ WITH (NOLOCK) WHERE TJ004=MB001 AND TJ001='A246' AND SUBSTRING(TJ002,1,6)=YEARMONTH) AS INT))  AS '實出量', ROUND((((SELECT ISNULL(SUM(TH008+TH024),0) FROM [{0}].dbo.COPTH WITH (NOLOCK) WHERE TH004=MB001 AND TH001='A233' AND SUBSTRING(TH002,1,6)=YEARMONTH)-(SELECT ISNULL(SUM(TJ007),0) FROM [{0}].dbo.COPTJ WITH (NOLOCK) WHERE TJ004=MB001 AND TJ001='A246' AND SUBSTRING(TJ002,1,6)=YEARMONTH)))/NULLIF([PREOrderNum],0)*100,2)  AS '銷售百分比' FROM [TKECOMMERCE].[dbo].[ZTKECOMMERCEFrmMPRECOPTC] WITH (NOLOCK) WHERE {1} ", NowDB, sbSqlQuery.ToString());
                NowTable = "TEMP1";
            }
            
            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            ds.Clear();

            if (comboBox1.Text.ToString().Equals("銷售完成率"))
            {
                adapter.Fill(ds, NowTable);
                dataGridView1.DataSource = ds.Tables[NowTable];
            }
            
            sqlConn.Close();
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExportExcel(ds, NowTable);
        }
        #endregion


    }
}
