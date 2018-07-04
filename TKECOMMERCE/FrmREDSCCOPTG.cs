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
    public partial class FrmREDSCCOPTG : Form
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
        DataSet ds= new DataSet();
        DataTable dt = new DataTable();
        string NowDB;
        string NowTable = null;
        int result;

        public FrmREDSCCOPTG()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void ExportExcel(DataSet dsExcel,string Tabelname)
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
            filename.AppendFormat(@"c:\temp\銷貨{0}.xlsx",DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-"+ filename.ToString());
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

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);



            sbSql.Clear();
            sbSqlQuery.Clear();

            if(checkBox1.Checked==true)
            {
                sbSqlQuery.Append("   (TG001='A233'  OR (TG001='A230'  AND TG006  IN ('160092','170007') )) AND ");
            }
            else if (checkBox1.Checked != true)
            {
                sbSqlQuery.Append("   (TG001='A233'  ) AND ");
            }

            sbSqlQuery.AppendFormat("  TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

            if (comboBox1.Text.ToString().Equals("銷售明細"))
            {
                sbSql.AppendFormat(" SELECT TG001 AS '單別',TG002 AS '單號',TG003 AS '日期',TG004 AS '客代',TG007 AS '客戶' ,TH004 AS '品號',TH005 AS '品名',TH008 AS '數量',TH024 AS '贈品',TH009 AS '單位',TH013 AS '金額' FROM [{0}].dbo.COPTG,[{1}].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TH020='Y' AND {2} ", NowDB, NowDB, sbSqlQuery.ToString());
                NowTable = "TEMP1";
            }
            else if(comboBox1.Text.ToString().Equals("品號彙總"))
            {
                sbSql.AppendFormat(" SELECT TH004 AS '品號',TH005 AS '品名',CONVERT(real, SUM(TH008)) AS '數量',CONVERT(real, SUM(TH024)) AS '贈品',TH009 AS '單位',CONVERT(real, SUM(TH013)) AS '金額' FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TH020='Y' AND  {2}  GROUP BY TH004,TH005,TH009 ORDER BY SUM(TH008) DESC", NowDB, NowDB, sbSqlQuery.ToString());
                NowTable = "TEMP2";
            }
            else if(comboBox1.Text.ToString().Equals("金額日彙總"))
            {
                sbSql.AppendFormat(" SELECT TG003 AS '日期',CONVERT(real, SUM(TH008)) AS '數量',CONVERT(real, SUM(TH024)) AS '贈品',CONVERT(real, SUM(TH013)) AS '金額' FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TH020='Y' AND   {2}   GROUP BY TG003 ", NowDB, NowDB, sbSqlQuery.ToString());
                NowTable = "TEMP3";
            }

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();            
            ds.Clear();

            if (comboBox1.Text.ToString().Equals("銷售明細"))
            {
                adapter.Fill(ds, NowTable);
                dataGridView1.DataSource = ds.Tables[NowTable];
            }
            else if (comboBox1.Text.ToString().Equals("品號彙總"))
            {
                adapter.Fill(ds, NowTable);
                dataGridView1.DataSource = ds.Tables[NowTable];
            }
            else if (comboBox1.Text.ToString().Equals("金額日彙總"))
            {
                adapter.Fill(ds, NowTable);
                dataGridView1.DataSource = ds.Tables[NowTable];
            }

            sqlConn.Close();
        }
        #endregion

        #region BUTTON   

        private void button1_Click_1(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExportExcel(ds,NowTable);
        }

        #endregion


    }
}
