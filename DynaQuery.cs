
using System;
using System.Data;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Text;
using System.Xml;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;




namespace ST_cd27d51818884ac9aafc175c5337b11d.csproj
{
    [Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {

        #region VSTA generated code
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

       

      

        public void Main()
        {
            // TODO: Add your code here

            int i = 0;
            
            SqlConnection oConn = new SqlConnection();
            SqlCommand oCmd1 = new SqlCommand();
            SqlCommand oCmd2 = new SqlCommand();
            //string strConn = Dts.Variables["V_BT_CONN_STRING"].Value.ToString();
            string strConn = Dts.Connections["GBSIS BTSQLDB"].ConnectionString.ToString();
            string[] strConnArr = strConn.Split(Convert.ToChar(";"));
            strConn = strConnArr[0] + ";" + strConnArr[1] + ";" + strConnArr[3];
            oConn.ConnectionString = strConn;
            StringBuilder strQuery1 = new StringBuilder();
            StringBuilder strQuery2 = new StringBuilder();
            StringBuilder strDynamicQuery = new StringBuilder();
            strQuery1.Append("SELECT CECSE FROM dbo.XXAMD_AFE_SPEND_CATEGORY_D WITH (NOLOCK) WHERE IS_ACTIVE = 'Y'");
            strQuery2.Append("SELECT LBC FROM dbo.XXAMD_AFE_LOCATION_D WITH (NOLOCK) WHERE REGION = 'NA' AND SPEND_FLAG = 'Y'");
            oCmd1.CommandText = strQuery1.ToString();
            oCmd2.CommandText = strQuery2.ToString();
            SqlDataAdapter oSQLDataAdapter1 = new SqlDataAdapter();
            SqlDataAdapter oSQLDataAdapter2 = new SqlDataAdapter();
            oCmd1.Connection = oConn;
            oCmd2.Connection = oConn;
            oSQLDataAdapter1.SelectCommand = oCmd1;
            oSQLDataAdapter2.SelectCommand = oCmd2;
            DataTable dtGroup1 = new DataTable();
            DataTable dtGroup2 = new DataTable();
            oSQLDataAdapter1.Fill(dtGroup1);
            oSQLDataAdapter2.Fill(dtGroup2);

            strDynamicQuery.Append("select * ");
            strDynamicQuery.Append("from apps.XXGDW_OBI_GL_BALANCES_AFE_V f ");
            strDynamicQuery.Append("where ");
            strDynamicQuery.Append("gl_period_fk_key >= to_number(to_char(add_months(sysdate,-1),'YYYYMM')) ");
            strDynamicQuery.Append("and source_fk_key = 7  /*NAM*/ ");
            strDynamicQuery.Append("and lbc_name  in ( ");
            if ((dtGroup2.Rows.Count > 0))
            {
                for (i = 0; i <= dtGroup2.Rows.Count - 1; i++)
                {
                    if ((string.Compare(dtGroup2.Rows[i]["LBC"].ToString(), "") != 0))
                    {
                        strDynamicQuery.Append("'" + dtGroup2.Rows[i]["LBC"].ToString() + "',");
                    }
                }
            }
            strDynamicQuery.Remove(strDynamicQuery.Length - 1, 1);
            strDynamicQuery.Append(")");
            strDynamicQuery.Append("and cecse_name in ( ");
            if ((dtGroup1.Rows.Count > 0))
            {
                for (i = 0; i <= dtGroup1.Rows.Count - 1; i++)
                {
                    if ((string.Compare(dtGroup1.Rows[i]["CECSE"].ToString(), "") != 0))
                    {
                        strDynamicQuery.Append("'" + dtGroup1.Rows[i]["CECSE"].ToString() + "',");
                    }
                }
            }
            strDynamicQuery.Remove(strDynamicQuery.Length - 1, 1);
            strDynamicQuery.Append(")");
            
            Dts.Variables["V_NAM_DYNAMIC_QUERY"].Value = strDynamicQuery.ToString();
            Dts.TaskResult = (int)ScriptResults.Success;
        }
    }
}