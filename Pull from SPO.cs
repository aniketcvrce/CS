


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
using Microsoft.SharePoint.Client;
using System.Security;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;
#endregion

namespace ST_47209dd1c5e146c7b2fdd877c133fcc4
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
    [Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {
        

        
        public void Main()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            string URL = Dts.Variables["$Project::SP_URL"].Value.ToString();
            string userName = Dts.Variables["$Project::SP_UserID"].Value.ToString();
            string passwordString = Convert.ToString(Dts.Variables["$Project::SP_PW"].GetSensitiveValue());
            string ListName = Dts.Variables["$Package::List_Name"].Value.ToString();
            string ConnectionName = Dts.Variables["$Project::SQL_ADO_Conn"].Value.ToString();

            PullData(URL, userName, passwordString, ListName, ConnectionName);

        }

        private static void PullData(string URL, string userName, string passwordString, string ListName, string ConnectionName)
        {
            #region DataTable
            DataTable dt = new DataTable();

            dt.Columns.Add("ID", typeof(Int32));
            dt.Columns.Add("Title");
            dt.Columns.Add("Participants_Name");
            dt.Columns.Add("Department");
            dt.Columns.Add("Feedback_Data_Time", typeof(DateTime));
            dt.Columns.Add("Feedback_Submit_Name");
            dt.Columns.Add("Feedback_Submit_Email");
            dt.Columns.Add("Task_Description");
            dt.Columns.Add("Is_Task_Same");
            dt.Columns.Add("Is_Task_Understandable");
            dt.Columns.Add("Task_Error_Description");
            dt.Columns.Add("Hazards_Type");
            dt.Columns.Add("Risk_Perception");
            dt.Columns.Add("Risk_Tolerance");
            dt.Columns.Add("Stop_Criteria");
            dt.Columns.Add("Is_Task_Safe");
            dt.Columns.Add("Help_Chain");
            dt.Columns.Add("Is_Task_Performed");
            dt.Columns.Add("Locations");
            dt.Columns.Add("Leader");
            dt.Columns.Add("Created_By");
            dt.Columns.Add("Creation_Date", typeof(DateTime));
            dt.Columns.Add("Modified_By");
            dt.Columns.Add("Modified_Date", typeof(DateTime));



            #endregion

            using (var ctx = new ClientContext(URL))
            {
                

                var passWord = new SecureString();
                foreach (char c in passwordString.ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials(userName, passWord);
                // Actual code for operations
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();

                ListItemCollectionPosition position = null;
                var page = 1;
                string HT = string.Empty;

                do
                {
                    CamlQuery query = new CamlQuery();

                    query.ViewXml = @"<View Scope='Recursive'>
                                      <Query>
                                      </Query><RowLimit>5000</RowLimit>
                                  </View>";
                    query.ListItemCollectionPosition = position;


                    List myList = web.Lists.GetByTitle(ListName);
                    ListItemCollection kundeItems = myList.GetItems(query);
                    ctx.Load<List>(myList);
                    ctx.Load<ListItemCollection>(kundeItems);
                    ctx.ExecuteQuery();

                    position = kundeItems.ListItemCollectionPosition;

                    DataRow dr;

                    #region for each row loop
                    foreach (var spItem in kundeItems)
                    {
                        dr = dt.NewRow();

                        try { dr["ID"] = spItem["ID"]; }
                        catch (Exception e) { dr["ID"] = DBNull.Value; }

                        try { dr["Title"] = spItem["Title"]; }
                        catch (Exception e) { dr["Title"] = DBNull.Value; }

                        try { dr["Participants_Name"] = spItem["Participants_Name"]; }
                        catch (Exception e) { dr["Participants_Name"] = DBNull.Value; }

                        try { dr["Department"] = spItem["Department"]; }
                        catch (Exception e) { dr["Department"] = DBNull.Value; }

                        try { dr["Feedback_Data_Time"] = spItem["Feedback_Data_Time"]; }
                        catch (Exception e) { dr["Feedback_Data_Time"] = DBNull.Value; }

                        try { dr["Feedback_Submit_Name"] = spItem["Feedback_Submit_Name"]; }
                        catch (Exception e) { dr["Feedback_Submit_Name"] = DBNull.Value; }

                        try { dr["Feedback_Submit_Email"] = spItem["Feedback_Submit_Email"]; }
                        catch (Exception e) { dr["Feedback_Submit_Email"] = DBNull.Value; }

                        try { dr["Task_Description"] = spItem["Task_Description"]; }
                        catch (Exception e) { dr["Task_Description"] = DBNull.Value; }

                        try { dr["Is_Task_Same"] = spItem["Is_Task_Same"]; }
                        catch (Exception e) { dr["Is_Same_Task"] = DBNull.Value; }

                        try { dr["Is_Task_Understandable"] = spItem["Is_Task_Understandable"]; }
                        catch (Exception e) { dr["Is_Task_Understandable"] = DBNull.Value; }

                        try { dr["Task_Error_Description"] = spItem["Task_Error_Description"]; }
                        catch (Exception e) { dr["Task_Error_Description"] = DBNull.Value; }


                        try {  HT = ((string[])spItem["Hazards_Type"])[0]; }
                        catch (Exception e) { dr["Hazards_Type"] = DBNull.Value; }
                        try { HT = string.Concat(HT, " ",((string[])spItem["Hazards_Type"])[1]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[2]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[3]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[4]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[5]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[6]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[7]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[8]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }
                        try { HT = string.Concat(HT, " ", ((string[])spItem["Hazards_Type"])[9]); }
                        catch (Exception e) { dr["Hazards_Type"] = HT; }



                        try { dr["Risk_Perception"] = spItem["Risk_Perception"]; }
                        catch (Exception e) { dr["Risk_Perception"] = DBNull.Value; }

                        try { dr["Risk_Tolerance"] = spItem["Risk_Tolerance"]; }
                        catch (Exception e) { dr["Risk_Tolerance"] = DBNull.Value; }

                        try { dr["Stop_Criteria"] = spItem["Stop_Criteria"]; }
                        catch (Exception e) { dr["Stop_Criteria"] = DBNull.Value; }

                        try { dr["Is_Task_Safe"] = spItem["Is_Task_Safe"]; }
                        catch (Exception e) { dr["Is_Task_Safe"] = DBNull.Value; }

                        try { dr["Help_Chain"] = spItem["Help_Chain"]; }
                        catch (Exception e) { dr["Help_Chain"] = DBNull.Value; }

                        try { dr["Stop_Criteria"] = spItem["Stop_Criteria"]; }
                        catch (Exception e) { dr["Stop_Criteria"] = DBNull.Value; }

                        try { dr["Is_Task_Performed"] = spItem["Is_Task_Performed"]; }
                        catch (Exception e) { dr["Is_Task_Performed"] = DBNull.Value; }

                        try { dr["Locations"] = spItem["Location"]; }
                        catch (Exception e) { dr["Location"] = DBNull.Value; }

                        try { dr["Leader"] = spItem["Leader"]; }
                        catch (Exception e) { dr["Leader"] = DBNull.Value; }
 
                       try
                        {
                            if (Convert.ToDateTime(spItem["Modified"]).Year == 1)
                                dr["Modified_Date"] = DBNull.Value;
                            else
                                dr["Modified_Date"] = System.DateTime.Parse(Convert.ToString(spItem["Modified"]), new CultureInfo("en-US", true));
                        }
                        catch (Exception e) { dr["Modified_Date"] = DBNull.Value; }
                        try
                        {
                            if (Convert.ToDateTime(spItem["Created"]).Year == 1)
                                dr["Creation_Date"] = DBNull.Value;
                            else
                                dr["Creation_Date"] = System.DateTime.Parse(Convert.ToString(spItem["Created"]), new CultureInfo("en-US", true));
                        }
                        catch (Exception e) { dr["Creation_Date"] = DBNull.Value; }

                        try
                        {
                            Microsoft.SharePoint.Client.FieldUserValue itemValue3 = (Microsoft.SharePoint.Client.FieldUserValue)spItem["Author"];
                            if (itemValue3.LookupValue == "service_spmigration2")
                                dr["Created_By"] = "";
                            else if (itemValue3.LookupValue == "service_spmigration2@arconic.com")
                                dr["Created_By"] = "";
                            else
                                dr["Created_By"] = itemValue3.LookupValue;
                        }
                        catch (Exception e) { dr["Created_By"] = ""; }

                        try
                        {
                            Microsoft.SharePoint.Client.FieldUserValue itemValue4 = (Microsoft.SharePoint.Client.FieldUserValue)spItem["Editor"];
                            if (itemValue4.LookupValue == "service_spmigration2")
                                dr["Modified_By"] = "";
                            else if (itemValue4.LookupValue == "service_spmigration2@arconic.com")
                                dr["Modified_By"] = "";
                            else
                                dr["Modified_By"] = itemValue4.LookupValue;
                        }
                        catch (Exception e) { dr["Modified_By"] = ""; }


                        dt.Rows.Add(dr);
                    }
                    #endregion

                    page++;


                }
                while (position != null);

                using (var bulkCopy = new SqlBulkCopy(ConnectionName, SqlBulkCopyOptions.KeepIdentity))
                {
                    

                    bulkCopy.ColumnMappings.Add("ID", "ID");
                    bulkCopy.ColumnMappings.Add("Title", "Title");
                    bulkCopy.ColumnMappings.Add("Participants_Name", "Participants_Name");
                    bulkCopy.ColumnMappings.Add("Department", "Department");
                    bulkCopy.ColumnMappings.Add("Feedback_Data_Time", "Feedback_Data_Time");
                    bulkCopy.ColumnMappings.Add("Feedback_Submit_Name", "Feedback_Submit_Name");
                    bulkCopy.ColumnMappings.Add("Feedback_Submit_Email", "Feedback_Submit_Email");
                    bulkCopy.ColumnMappings.Add("Is_Task_Same", "Is_Task_Same");
                    bulkCopy.ColumnMappings.Add("Is_Task_Understandable", "Is_Task_Understandable");
                    bulkCopy.ColumnMappings.Add("Task_Error_Description", "Task_Error_Description");
                    bulkCopy.ColumnMappings.Add("Hazards_Type", "Hazards_Type");
                    bulkCopy.ColumnMappings.Add("Risk_Perception", "Risk_Perception");
                    bulkCopy.ColumnMappings.Add("Risk_Tolerance", "Risk_Tolerance");
                    bulkCopy.ColumnMappings.Add("Stop_Criteria", "Stop_Criteria");
                    bulkCopy.ColumnMappings.Add("Is_Task_Safe", "Is_Task_Safe");
                    bulkCopy.ColumnMappings.Add("Help_Chain", "Help_Chain");
                    bulkCopy.ColumnMappings.Add("Is_Task_Performed", "Is_Task_Performed");
                    bulkCopy.ColumnMappings.Add("Locations", "Locations");
                    bulkCopy.ColumnMappings.Add("Leader", "Leader");
                    bulkCopy.ColumnMappings.Add("Created_By", "Created_By");
                    bulkCopy.ColumnMappings.Add("Creation_Date", "Creation_Date");
                    bulkCopy.ColumnMappings.Add("Modified_By", "Modified_By");
                    bulkCopy.ColumnMappings.Add("Modified_Date", "Modified_Date");


                    bulkCopy.BulkCopyTimeout = 600;
                    bulkCopy.DestinationTableName = "Am_I_Ready_STG";
                    bulkCopy.WriteToServer(dt);
                }



            }


        }



    }
}