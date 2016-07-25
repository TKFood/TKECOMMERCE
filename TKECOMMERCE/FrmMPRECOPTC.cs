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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSqlQuery.AppendFormat("{0}",dateTimePicker1.Value.ToString("yyyyMM"));
                    sbSql.AppendFormat(@"SELECT  YEARMONTH AS '年月', ZTKECOMMERCEFrmMPRECOPTC.MB001 AS '品號', ZTKECOMMERCEFrmMPRECOPTC.MB002 AS '品名', PREOrderNum AS '數量',MB004 AS '單位' ,TC001 AS '訂單別' ,TC002 AS '訂單號' FROM  [{0}].[dbo].ZTKECOMMERCEFrmMPRECOPTC  WITH (NOLOCK) ,[{1}].[dbo].INVMB  WITH (NOLOCK) WHERE ZTKECOMMERCEFrmMPRECOPTC.MB001=INVMB.MB001 AND YEARMONTH='{2}'  ", sqlConn.Database.ToString(), NowDB, sbSqlQuery.ToString());

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
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" ");               
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sbSql.AppendFormat(" INSERT INTO [{0}].[dbo].[ZTKECOMMERCEFrmMPRECOPTC] ", sqlConn.Database.ToString());
                    sbSql.Append(" (YEARMONTH, MB001,MB002,  PREOrderNum )");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','NA','{2}')", dt.Rows[i]["年月"].ToString(), dt.Rows[i]["品號"].ToString(), dt.Rows[i]["數量"].ToString());
                    sbSql.Append(" ");
                }
                sbSql.AppendFormat(" UPDATE [{0}].[dbo].[ZTKECOMMERCEFrmMPRECOPTC] SET ZTKECOMMERCEFrmMPRECOPTC.MB002=INVMB.MB002 FROM [{1}].[dbo].[INVMB] WHERE ZTKECOMMERCEFrmMPRECOPTC.MB001=INVMB.MB001 AND ZTKECOMMERCEFrmMPRECOPTC.MB002='NA'", sqlConn.Database.ToString(), NowDB);
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
                sqlConn.Close();

                
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

        public void AddtoERP()
        {
            string TC001 = "A223";
            string TC002 ="0";
            USEDFUNCTION FUNGetMaxID = new USEDFUNCTION();

            TC002 = FUNGetMaxID.GetMaxID(TC001);

            try
            {
                if(ds.Tables["TEMPds"].Rows.Count>=1)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //ADD COPTC
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" INSERT INTO [{0}].[dbo].COPTC", NowDB);
                    sbSql.Append(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.Append(" ,[TC001],[TC002],[TC003],[TC004],[TC005],[TC006],[TC007],[TC008],[TC009],[TC010]");
                    sbSql.Append(" ,[TC011],[TC012],[TC013],[TC014],[TC015],[TC016],[TC017],[TC018],[TC019],[TC020]");
                    sbSql.Append(" ,[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030]");
                    sbSql.Append(" ,[TC031],[TC032],[TC033],[TC034],[TC035],[TC036],[TC037],[TC038],[TC039],[TC040]");
                    sbSql.Append(" ,[TC041],[TC042],[TC043],[TC044],[TC045],[TC046],[TC047],[TC048],[TC049],[TC050]");
                    sbSql.Append(" ,[TC051],[TC052],[TC053],[TC054],[TC055],[TC056],[TC057],[TC058],[TC059],[TC060] ");
                    sbSql.Append(" ,[TC061],[TC062],[TC063],[TC064],[TC065],[TC066],[TC067],[TC068],[TC069],[TC070] ");
                    sbSql.Append(" ,[TC071],[TC072],[TC073],[TC074],[TC075],[TC076],[TC077],[TC078],[TC079],[TC080]");
                    sbSql.Append(" ,[TC081],[TC082],[TC083],[TC084],[TC085],[TC086],[TC087],[TC088],[TC089],[TC090]");
                    sbSql.Append(" ,[TC091],[TC092],[TC093],[TC094],[TC095],[TC096],[TC097],[TC098],[TC099],[TC100]");
                    sbSql.Append(" ,[TC101],[TC102],[TC103],[TC104],[TC105],[TC106],[TC107])");
                    sbSql.Append(" VALUES");
                    sbSql.AppendFormat(" ('TK' , 'DS' , 'DS',SUBSTRING('{0}',1,8) , NULL , NULL , '0' ,'12:00:01', NULL, 'P001', 'COPMI06', NULL, NULL , NULL , '0' , 'DS', 'DS'", TC002);
                    sbSql.AppendFormat(" , 'A223' , '{0}' , '{1}' , '91000005' , '106400' , '160092' , '20' , 'NTD' , 1 , NULL ", TC002, DateTime.Now.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(" , NULL , NULL , NULL , NULL , '電子商務{0}月預估單' , '1' , NULL , NULL , '1' , NULL ", TC002.Substring(0, 6));
                    sbSql.Append(" , NULL , NULL , NULL , NULL , NULL , 0 , 'N' , 0 , 0 , 0 ");
                    sbSql.AppendFormat(" , 0 , '91000005' , NULL , NULL , NULL , NULL , NULL , NULL , SUBSTRING('{0}',1,8) , NULL ", TC002);
                    sbSql.Append(" , 0.05 , NULL , 0 , 0 , 0 , 0 , NULL , 0 , NULL , 'N' ");
                    sbSql.Append(" , NULL , 0 , '電子商務訂單客戶' , NULL , NULL , NULL , NULL , NULL , NULL , NULL ");
                    sbSql.Append(" , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL ");
                    sbSql.Append(" , NULL , NULL , 0 , NULL , NULL , NULL , NULL , NULL , NULL , NULL ");
                    sbSql.Append(" , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL ");
                    sbSql.Append(" , NULL , 'N' , NULL , NULL , NULL , NULL , NULL , NULL , NULL , NULL  ");
                    sbSql.Append(", NULL , NULL , NULL , NULL , NULL , NULL , NULL  ) ");
                    sbSql.Append(" ");
                    //ADD COPTD
  
                    sbSql.Append(" INSERT INTO [test].[dbo].[COPTD]");
                    sbSql.Append(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.Append(" ,[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010]");
                    sbSql.Append(" ,[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020]");
                    sbSql.Append(" ,[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD029],[TD030]");
                    sbSql.Append(" ,[TD031],[TD032],[TD033],[TD034],[TD035],[TD036],[TD037],[TD038],[TD039],[TD040]");
                    sbSql.Append(" ,[TD041],[TD042],[TD043],[TD044],[TD045],[TD046],[TD047],[TD048],[TD049],[TD050]");
                    sbSql.Append(" ,[TD051],[TD052],[TD053],[TD054],[TD055],[TD056],[TD057],[TD058],[TD059],[TD060]");
                    sbSql.Append(" ,[TD061],[TD062],[TD063],[TD064],[TD065],[TD066],[TD067],[TD068],[TD069],[TD070]");
                    sbSql.Append(" ,[TD071],[TD072],[TD073],[TD074],[TD075],[TD076],[TD077],[TD078],[TD079],[TD080]");
                    sbSql.Append(" ,[TD081],[TD082],[TD083],[TD084],[TD085],[TD086],[TD087])");
                    sbSql.Append(" SELECT ");
                    sbSql.Append(" 'TK' AS [COMPANY], 'DS' AS [CREATOR], 'DS' AS  [USR_GROUP], SUBSTRING('20160721',1,8) [CREATE_DATE], NULL [MODIFIER], NULL [MODI_DATE], '0' [FLAG], '12:00:01' [CREATE_TIME], NULL [MODI_TIME], 'P001' [TRANS_TYPE], 'COPMI06' [TRANS_NAME], NULL [sync_date], NULL [sync_time], NULL [sync_mark], '0' [sync_count], 'DS' [DataUser], 'DS' [DataGroup]");
                    sbSql.AppendFormat(" , '{0}' [TD001], '{1}'  [TD002],  RIGHT('0000' + CAST(row_number() OVER(PARTITION BY YEARMONTH ORDER BY YEARMONTH) as varchar), 4)   [TD003], ZTKECOMMERCEFrmMPRECOPTC.MB001  [TD004],ZTKECOMMERCEFrmMPRECOPTC.MB002 [TD005], MB003 [TD006], '20001' [TD007], PREOrderNum [TD008], 0 [TD009], MB004 [TD010]", TC001.ToString(),TC002.ToString());
                    sbSql.AppendFormat(" , MB047 [TD011], PREOrderNum*MB047 [TD012], {0} [TD013], NULL [TD014], NULL [TD015], 'N' [TD016], NULL [TD017], NULL [TD018], NULL [TD019], NULL [TD020]",TC002.Substring(0,8));
                    sbSql.Append(" , 'N' [TD021], 0 [TD022], NULL [TD023], 0 [TD024], 0 [TD025], 1 [TD026], NULL [TD027], NULL [TD028], NULL [TD029], 0 [TD030]");
                    sbSql.Append(" , 0 [TD031], 0 [TD032], 0 [TD033], 0 [TD034], 0 [TD035], NULL [TD036], NULL [TD037], NULL [TD038], NULL [TD039], NULL [TD040]");
                    sbSql.Append(" , NULL [TD041], 0 [TD042], NULL [TD043], NULL [TD044], '9' [TD045], NULL [TD046], NULL [TD047], NULL [TD048], '1' [TD049], 0 [TD050]");
                    sbSql.Append(" , 0 [TD051], 0 [TD052], 0 [TD053], 0 [TD054], 0 [TD055], NULL [TD056], NULL [TD057], NULL [TD058], 0 [TD059], '1' [TD060]");
                    sbSql.Append("   , 0 [TD061], 'N' [TD062], NULL [TD063], NULL [TD064], NULL [TD065], NULL [TD066], NULL [TD067], NULL [TD068],'N'  [TD069], 0 [TD070] ");
                    sbSql.Append(" , NULL [TD071], NULL [TD072], NULL [TD073], NULL [TD074], NULL [TD075], 0 [TD076], NULL [TD077], NULL [TD078],'N' [TD079], NULL [TD080]");
                    sbSql.Append(" , 0 [TD081], NULL [TD082], NULL [TD083], NULL [TD084], 0 [TD085], NULL [TD086], 0 [TD087]");
                    sbSql.Append(" FROM [TKECOMMERCE].[dbo].[ZTKECOMMERCEFrmMPRECOPTC],[test].[dbo].[INVMB]");
                    sbSql.Append(" WHERE ZTKECOMMERCEFrmMPRECOPTC.MB001=INVMB.MB001");
                    sbSql.AppendFormat(" AND YEARMONTH='{0}'",dateTimePicker1.Value.ToString("yyyyMM"));
                    sbSql.Append(" ");
                    //UPDATE
                    sbSql.AppendFormat(" UPDATE [{0}].dbo.COPTC SET ", NowDB);
                    sbSql.AppendFormat(" TC029=(SELECT ROUND(SUM(TD012)/1.05,0) FROM  [{0}].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002)", NowDB);
                    sbSql.AppendFormat(" ,TC030=(SELECT SUM(TD012)-ROUND(SUM(TD012)/1.05,0) FROM  [{0}].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002)", NowDB);
                    sbSql.AppendFormat(" ,TC031=(SELECT SUM(TD008) FROM  [{0}].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002)", NowDB);
                    sbSql.AppendFormat(" WHERE TC001='{0}' AND TC002='{1}'", TC001, TC002);

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
                        textBox2.Text = TC001.ToString();
                        textBox3.Text = TC002.ToString();
                    }

                    sqlConn.Close();

                    //UPDATE ZTKECOMMERCEFrmMPRECOPTC
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //ADD COPTC
                    sbSql.AppendFormat(" UPDATE [{0}].dbo.[ZTKECOMMERCEFrmMPRECOPTC] SET TC001='{1}',TC002='{2}' WHERE YEARMONTH='{3}'",NowDB ,TC001.ToString(),TC002.ToString(),dateTimePicker1.Value.ToString("yyyyMM"));

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

                    sqlConn.Close();
                }
                

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            Search();

        }

        public void DelZTKECOMMERCEFrmMPRECOPTC()
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("是否真的要刪除", "del?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //Del ZTKECOMMERCEFrmMPRECOPTC
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //ADD COPTC
                    sbSql.AppendFormat(" DELETE [{0}].dbo.[ZTKECOMMERCEFrmMPRECOPTC] WHERE YEARMONTH='{1}'", NowDB, dateTimePicker1.Value.ToString("yyyyMM"));

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

                    sqlConn.Close();
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }

                
            }
            catch
            {

            }
            finally
            {

            }

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
        private void button4_Click(object sender, EventArgs e)
        {
            AddtoERP();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DelZTKECOMMERCEFrmMPRECOPTC();
        }
        #endregion


    }
}
