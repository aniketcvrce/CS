#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text;
using Microsoft.AnalysisServices.AdomdClient;
#endregion

namespace ST_8627d4aeb3de49b488e44abd1323882a
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion


		/// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
		public void Main()
		{
            // TODO: Add your code here

            //$Package::ESSBASE_ACTUAL_SCENARIO,$Package::ESSBASE_ACTUAL_YEAR,$Package::ESSBASE_CURR_FCST_SCENARIO,$Package::ESSBASE_CURR_FCST_YEAR,$Package::ESSBASE_FCST_ARCHIVE_SCENARIO,$Package::ESSBASE_FCST_ARCHIVE_YEAR,$Package::ESSBASE_PLAN_SCENARIO,$Package::ESSBASE_PLAN_YEAR,$Package::EXECUTION_ID,$Package::SQL_SERVER_ADO_CONNSTR,$Project::ESSBASE_APPLICATION_NAME,$Project::ESSBASE_CUBE_NAME,$Project::ESSBASE_DATA_SOURCE_INFO,$Project::ESSBASE_PWD,$Project::ESSBASE_SERVER,$Project::ESSBASE_USER_ID

            #region Get Variables/Parameters

            string pw = Convert.ToString(Dts.Variables["$Project::SQL_SERVER_PW"].GetSensitiveValue());

            string planScenario = Convert.ToString(Dts.Variables["$Package::ESSBASE_PLAN_SCENARIO"].Value);
            string planYear = Convert.ToString(Dts.Variables["$Package::ESSBASE_PLAN_YEAR"].Value);

            string actualScenario = Convert.ToString(Dts.Variables["$Package::ESSBASE_ACTUAL_SCENARIO"].Value);
            string actualYear = Convert.ToString(Dts.Variables["$Package::ESSBASE_ACTUAL_YEAR"].Value);

            string currFcstScenario = Convert.ToString(Dts.Variables["$Package::ESSBASE_CURR_FCST_SCENARIO"].Value);
            string currFcstYear = Convert.ToString(Dts.Variables["$Package::ESSBASE_CURR_FCST_YEAR"].Value);

            string fcstArchiveScenario = Convert.ToString(Dts.Variables["$Package::ESSBASE_FCST_ARCHIVE_SCENARIO"].Value);
            string fcstArchiveYear = Convert.ToString(Dts.Variables["$Package::ESSBASE_FCST_ARCHIVE_YEAR"].Value);

            string sqlConnStr = Convert.ToString(Dts.Variables["$Project::SQL_SERVER_ADO_CONNSTR"].Value) + "Password=" + pw + ";";
            string applicationName = Convert.ToString(Dts.Variables["$Project::ESSBASE_APPLICATION_NAME"].Value);
            string cubeName = Convert.ToString(Dts.Variables["$Project::ESSBASE_CUBE_NAME"].Value);

            string paramScenario = string.Empty, paramYear = string.Empty;

            if (!planScenario.Equals(string.Empty) && !planYear.Equals(string.Empty))
            {
                paramScenario = planScenario;
                paramYear = planYear;
            }
            else if (!actualScenario.Equals(string.Empty) && !actualYear.Equals(string.Empty))
            {
                paramScenario = actualScenario;
                paramYear = actualYear;
            }
            else if (!currFcstScenario.Equals(string.Empty) && !currFcstYear.Equals(string.Empty))
            {
                paramScenario = currFcstScenario;
                paramYear = currFcstYear;
            }
            else if (!fcstArchiveScenario.Equals(string.Empty) && !fcstArchiveYear.Equals(string.Empty))
            {
                paramScenario = fcstArchiveScenario;
                paramYear = fcstArchiveYear;
            }
            else
            { 
                // Log error - all parameter values not supplied.
                WriteLog("Either Scenario or Year or both are not supplied to Child package.");
                WriteLog("planScenario:" + planScenario + ";planYear:" + planYear + ";actualScenario:" + actualScenario + ";actualYear:" + actualYear +
                    ";currFcstScenario:" + currFcstScenario + ";currFcstYear:" + currFcstYear + ";fcstArchieveScenario:" + fcstArchiveScenario + ";fcstArchieveYear:" + fcstArchiveYear);
            }

            #endregion

            

            string[] months = { "BegBalance", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

            if (paramScenario.Equals(string.Empty) || paramYear.Equals(string.Empty))
            {
                Dts.TaskResult = (int)ScriptResults.Failure;
            }
            else
            {
                WriteLog("Parameters from Master Package - paramScenario:" + paramScenario + ";paramYear:" + paramYear);

                #region Get Entities and Accounts from SQL Table

                DataTable dtEntities = new DataTable();
                DataTable dtAccounts = new DataTable();

                SqlConnection sqlConn = new SqlConnection(sqlConnStr);
                string sqlGetEntity = "SELECT ENTITY_NAME FROM XXBUKPI_ENTITIES_SSTG WHERE IS_ACTIVE = 1";
                string sqlGetAccount = "SELECT ACCOUNT_NAME FROM XXBUKPI_ACCOUNT_SSTG WHERE IS_ACTIVE = 1";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(sqlGetEntity, sqlConn);
                dataAdapter.Fill(dtEntities);

                dataAdapter = new SqlDataAdapter(sqlGetAccount, sqlConn);
                dataAdapter.Fill(dtAccounts);

                string[] scenario = paramScenario.Split('|');
                //SCENARIOS = Convert.ToString(Dts.Variables["$Package::ESSBASE_SCENARIO"].Value).Split('|');

                string[] year = paramYear.Split('|');

                //foreach (string scenario in SCENARIOS)
                for (int i = 0; i < scenario.Length; i++)
                {
                    string[] years = year[i].Split(',');
                    string pYear = string.Empty;

                    foreach (string item in years)
                    {
                        // pYear = pYear + "[" + item + "]" + ",";
                        pYear = "[" + item + "]";
                        //}

                        pYear = pYear.TrimEnd(',');

                        foreach (DataRow row in dtEntities.Rows)
                        {
                            string entities = string.Empty;
                            entities += "Descendants(";
                            entities += Convert.ToString(row["ENTITY_NAME"]);

                            entities += ")";

                            foreach (DataRow dtrow in dtAccounts.Rows)
                            {
                                string accounts = string.Empty;
                                accounts += "Descendants(";
                                accounts += Convert.ToString(dtrow["ACCOUNT_NAME"]);
                                accounts += ")";

                                for (int m = 0; m <= 12; m++)
                                {

                                    WriteLog("Scenario:" + scenario[i] + ";Year:" + pYear + ";Entity:" + entities + ";Account:" + accounts + ";Month:" + months[m]);
                                    #endregion

                                    #region Formulate MDX query and retrieve data

                                    StringBuilder mdxQuery = new StringBuilder();

                                    mdxQuery.Append("SELECT {[");
                                    mdxQuery.Append(scenario[i]);
                                    mdxQuery.Append("]} ON COLUMNS, ");
                                    mdxQuery.Append(" NON EMPTY (CROSSJOIN(CROSSJOIN(CROSSJOIN (CROSSJOIN(CROSSJOIN(CROSSJOIN(CROSSJOIN(CROSSJOIN(");
                                    mdxQuery.Append(" {[USD],[LOCAL]}, ");
                                    mdxQuery.Append(" {");
                                    mdxQuery.Append(entities);
                                    mdxQuery.Append("}), ");
                                    mdxQuery.Append(" {");
                                    mdxQuery.Append(accounts.TrimEnd(','));
                                    mdxQuery.Append("}),{Descendants(CEC_A000)}), ");
                                    mdxQuery.Append(" {");
                                    mdxQuery.Append(pYear);
                                    mdxQuery.Append("}), ");
                                    //mdxQuery.Append(" {BegBalance,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec}),{[HSP_InputValue]}),");
                                    mdxQuery.Append(" {" + months[m] + "}),{[HSP_InputValue]}),");
                                    mdxQuery.Append("{[working]}),{[data_total]}) ");
                                    mdxQuery.Append(" ) DIMENSION PROPERTIES ");
                                    mdxQuery.Append(" [Entities].[MEMBER_NAME], ");
                                    mdxQuery.Append(" [Accounts].[MEMBER_NAME], ");
                                    mdxQuery.Append(" [Currencies].[MEMBER_NAME], ");
                                    mdxQuery.Append(" [Years].[MEMBER_NAME], ");
                                    mdxQuery.Append(" [Time Periods].[MEMBER_NAME],[CECSE].[MEMBER_NAME]   ");
                                    mdxQuery.Append(" ON ROWS ");
                                    mdxQuery.Append(" FROM ");
                                    mdxQuery.Append(applicationName);
                                    mdxQuery.Append(".");
                                    mdxQuery.Append(cubeName);
                                    //  mdxQuery.Append(" WHERE ([CECSE],[Working],[HSP_InputValue],[DATA_TOTAL])");                            
                                    //MessageBox.Show(mdxQuery.ToString());
                                    string connectionString = string.Empty;
                                    connectionString = "Data Source=http://" + Convert.ToString(Dts.Variables["$Project::ESSBASE_SERVER"].Value) + "/aps/XMLA; Initial Catalog=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_APPLICATION_NAME"].Value) + ";User Id=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_USER_ID"].Value) + ";Password=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_PWD"].GetSensitiveValue()) + ";DataSourceInfo=\"Provider=Essbase;Data Source=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_DATA_SOURCE_INFO"].Value) + "\";";
                                    //connectionString = "Provider=MSOLAP;Data Source=http://" + Convert.ToString(Dts.Variables["$Project::ESSBASE_SERVER"].Value) + "/aps/XMLA; Initial Catalog=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_APPLICATION_NAME"].Value) + ";User Id=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_USER_ID"].Value) + ";Password=" + Convert.ToString(Dts.Variables["$Project::ESSBASE_PWD"].GetSensitiveValue()) + ";";

                                    DataTable dtBUKPI = RetrieveEssBase(mdxQuery.ToString(), connectionString);

                                    WriteLog("Data retrieval complete for " + "Scenario:" + scenario[i] + ";Year:" + pYear + ";Entity:" + entities + ";Account:" + accounts + ";Month:" + months[m]);

                                    DataTable dtBUKPIFinal = FormatEssBaseData(dtBUKPI, scenario[i]);

                                    WriteLog("Data formatting complete for " + "Scenario:" + scenario[i] + ";Year:" + pYear + ";Entity:" + entities + ";Account:" + accounts + ";Month:" + months[m]);

                                    //Clear record for scenario-year combination and set destination table name for bulkcopy based on scenario

                                    string dstSQLTable = string.Empty;

                                    if (dtBUKPIFinal.Rows.Count > 0)
                                    {
                                        string sqlClearData = string.Empty;

                                        if (scenario[i].ToUpper().Equals("ACTUAL"))
                                            dstSQLTable = "BUKPI_EXTRACTION_ACTUAL_STG";
                                        else if (scenario[i].ToUpper().Equals("CURR FCST"))
                                            dstSQLTable = "BUKPI_EXTRACTION_CURR_FCST_STG";
                                        else if (scenario[i].ToUpper().Contains("FCST"))
                                            dstSQLTable = "BUKPI_EXTRACTION_FCST_ARCHIVE_STG";
                                        else if (scenario[i].ToUpper().Contains("PLAN"))
                                            dstSQLTable = "BUKPI_EXTRACTION_PLAN_STG";
                                        else
                                            dstSQLTable = "";

                                        if (dstSQLTable.Equals(""))
                                        {
                                            // Log Error. Scenario NOT handled.
                                            WriteLog("Scenario NOT handled; Scenario:" + scenario[i]);
                                        }
                                        else
                                        {
                                            using (var bulkCopy = new SqlBulkCopy(sqlConnStr, SqlBulkCopyOptions.KeepIdentity))
                                            {
                                                // my DataTable column names match my SQL Column names, so I simply made this loop. However if your column names don't match, just pass in which datatable name matches the SQL column name in Column Mappings
                                                foreach (DataColumn col in dtBUKPIFinal.Columns)
                                                {
                                                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                                                }

                                                bulkCopy.BulkCopyTimeout = 600;
                                                bulkCopy.DestinationTableName = dstSQLTable;
                                                bulkCopy.WriteToServer(dtBUKPIFinal);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Log Error. NO record returned from Hyperion
                                        WriteLog("NO record returned from Essbase Cube. MDX Query:" + mdxQuery);
                                    }
                                }
                            }
                        }
                    }
                }

                //Dts.Variables["User::V_BUKPI_INV_PROD"].Value = dtBUKPIInvProdFinal;
                #endregion

                Dts.TaskResult = (int)ScriptResults.Success;
            }
		}

        private DataTable FormatEssBaseData(DataTable dtBUKPIInvProd,string scenario)
        {
            DataTable dtFormattedBUKPIInvProd = new DataTable();
            dtFormattedBUKPIInvProd.Columns.Add("CURRENCY_CODE", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("ENTITY", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("ACCOUNT", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("CECSE", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("DATATYPE", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("VERSION", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("HSP_RATES", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("PERIOD", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("MONTH_NUMBER", typeof(int));
            dtFormattedBUKPIInvProd.Columns.Add("YEAR", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("CALENDAR_YEAR", typeof(int));
            dtFormattedBUKPIInvProd.Columns.Add("SCENARIO", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("VALUE", typeof(string));
            dtFormattedBUKPIInvProd.Columns.Add("LOAD_DATE", typeof(DateTime));
            
            //GET THE LAST AVAILABLE LEVEL OF EACH DIMENSIONS IN THE DATA TABLE CONTAINING ESSBASE DATA
            foreach (DataRow dr in dtBUKPIInvProd.Rows)
            {
                int i_Currencies = GetLastLevel("Currencies", dtBUKPIInvProd, dr);
                int i_Entities = GetLastLevel("Entities", dtBUKPIInvProd, dr);
                int i_Time_Periods = GetLastLevel("Time Periods", dtBUKPIInvProd, dr);
                int i_Years = GetLastLevel("Years", dtBUKPIInvProd, dr);
                int i_Accounts = GetLastLevel("Accounts", dtBUKPIInvProd, dr);
                int i_cecse = GetLastLevel("CECSE", dtBUKPIInvProd, dr);

                dtFormattedBUKPIInvProd.Rows.Add
                (
                 dr[i_Currencies],
                 dr[i_Entities],
                 dr[i_Accounts],
                 dr[i_cecse],
                 dr["[Datatype].Levels(1).[MEMBER_CAPTION]"],
                 dr["[Versions].Levels(1).[MEMBER_CAPTION]"],
                 dr["[HSP_Rates].Levels(1).[MEMBER_CAPTION]"],
                 dr[i_Time_Periods],
                 GetMonthNumber(Convert.ToString(dr[i_Time_Periods])),
                 dr[i_Years],
                 Convert.ToInt32("20" + Convert.ToString(dr[i_Years]).Substring(2)),
                 scenario,
                 dr["[" + scenario + "]"], //Value field
                 DateTime.Now
                );
            }
            
            return dtFormattedBUKPIInvProd;
        }

        private int GetLastLevel(string dimension_name, DataTable dtBUKPIInvProd, DataRow dr)
        {
            //start and end of the dimension level ordinals
            int i_start = -1;
            int i_end = 0;
            foreach (DataColumn dc in dtBUKPIInvProd.Columns)
            {
                if (dc.ColumnName.Contains(dimension_name))
                {
                    if (i_start == -1)
                        i_start = dc.Ordinal;
                    i_end = dc.Ordinal;
                }
            }
            int j;
            for (j = i_start; j <= i_end; j++)
            {
                if (Convert.ToString(dr.ItemArray[j]) == ""  )
                {
                    break;
                }
            }
            if (j != i_start)
                return j - 1;
            else return j;
        }

        private DataTable RetrieveEssBase(string query, string connectionString)
        {
            AdomdDataReader reader = null;
            AdomdConnection conn = new AdomdConnection(connectionString);
            DataTable EssBaseTable = new DataTable();
            try
            {
                conn.Open();
                
                WriteLog("Essbase connection successful. Starting data retrieval process.");

                AdomdCommand cmd = new AdomdCommand(query, conn);
                reader = cmd.ExecuteReader();

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    if (reader.GetName(i) == "MEMBER_NAME")
                        continue;
                    else
                        EssBaseTable.Columns.Add(reader.GetName(i), typeof(String));
                }
                while (reader.Read())
                {
                    object[] array = new object[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        array[i] = reader.GetValue(i);
                    }
                    EssBaseTable.Rows.Add(array);
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                WriteLog("Exception invoked from RetrieveEssBase method; Error Message - " + ex.Message);
                throw ex;
            }
            finally
            {
                //reader.Close();
                // conn.Close();
                //conn = null;

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            return EssBaseTable;
        }

        private int GetMonthNumber(string monthName)
        {
            int monthNumber;

            switch (monthName.ToUpper())
            { 
                case "JAN":
                    monthNumber = 1;
                    break;
                case "FEB":
                    monthNumber = 2;
                    break;
                case "MAR":
                    monthNumber = 3;
                    break;
                case "APR":
                    monthNumber = 4;
                    break;
                case "MAY":
                    monthNumber = 5;
                    break;
                case "JUN":
                    monthNumber = 6;
                    break;
                case "JUL":
                    monthNumber = 7;
                    break;
                case "AUG":
                    monthNumber = 8;
                    break;
                case "SEP":
                    monthNumber = 9;
                    break;
                case "OCT":
                    monthNumber = 10;
                    break;
                case "NOV":
                    monthNumber = 11;
                    break;
                case "DEC":
                    monthNumber = 12;
                    break;
                case "BEGBALANCE":
                    monthNumber = 13;
                    break;
                default:
                    monthNumber = 0;
                    break;
            }

            return monthNumber;
        }

        private void WriteLog(string logMsg)
        {
            string pw = Convert.ToString(Dts.Variables["$Project::SQL_SERVER_PW"].GetSensitiveValue());
            string sqlConnStr = Convert.ToString(Dts.Variables["$Project::SQL_SERVER_ADO_CONNSTR"].Value) + "Password=" + pw + ";";

            string executionID = Convert.ToString(Dts.Variables["$Package::EXECUTION_ID"].Value);
            string source = "XXBUKPI_LOAD_BUKPI_DATA_CHILD";
            string sql = "INSERT INTO [dbo].[XXBUKPI_LOG] (EXECUTION_ID, SOURCE, MESSAGE, LOGGED_DATE) VALUES ('" + executionID + "','" + source + "','" + logMsg + "','" + DateTime.Now + "')";

            //SqlConnection sqlConn = new SqlConnection(Convert.ToString(Dts.Variables["$Package::SQL_SERVER_ADO_CONNSTR"].Value));
            SqlConnection sqlConn = new SqlConnection(sqlConnStr);
            


            if (sqlConn.State == ConnectionState.Closed)
                sqlConn.Open();

            SqlCommand sqlCmd = new SqlCommand(sql, sqlConn);
            sqlCmd.ExecuteNonQuery();

            if (sqlConn.State == ConnectionState.Open)
                sqlConn.Close();
        }

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

	}
}