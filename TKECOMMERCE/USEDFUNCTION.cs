using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TKECOMMERCE
{
    class USEDFUNCTION
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

        public string GetMaxID(string TC001)
        {
            string newid;
            int countid;
            NowDay = DateTime.Now.ToString("yyyyMMdd");
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendFormat(@"SELECT( CASE WHEN ISNULL(MAX(TC002),'')='' THEN '0' ELSE  MAX(TC002)  END) AS TC002  FROM  [{2}].dbo.COPTC WITH (NOLOCK) WHERE TC003='{0}' AND TC001='{1}' ", NowDay, TC001, NowDB);

            DataSet dt = new DataSet();
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sbSql.ToString(), sqlConn);

            sqlConn.Open();
            adapter = new SqlDataAdapter(cmd);
            dt.Clear();
            adapter.Fill(dt);

            newid = dt.Tables[0].Rows[0][0].ToString();
            if (newid.ToString().Equals("0"))
            {
                countid = 0;
            }
            else
            {
                countid = Convert.ToInt16(newid.Substring(8, 3));
            }


            countid = countid + 1;
            newid = NowDay + countid.ToString().PadLeft(3, '0');

            return newid;
        }
    }
}
