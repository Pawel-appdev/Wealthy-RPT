
using CsvHelper;
using ExPat_UI.Reports;
using ExPat_UI.Utilities;
using Wealthy_RPT;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;

namespace ExPat_UI.Imports
{
    class ImportFunctions : Reports.ReportFunctions
    {
        SqlDataAdapter adapter = new SqlDataAdapter();
        public DataSet ds = new DataSet();

        /// <summary> Public Method ImportCSVFile
        /// method called from menus accepts destinaton name,
        /// sets corresponding DB tablename and starts Import process.
        /// </summary>
        /// <param name="importTable">Destination Name</param>
        /// <returns>Bool of success/failure</returns>
        public static bool ImportCSVFile(string importTable)
        {
            string dbTableName = null;
            switch (importTable)
            {
                case "CaseflowHistory":
                    dbTableName = "[tblCaseHistory]";
                    break;
                case "UpstreamHistory":
                    dbTableName = "[tblUpstreamHistory]";
                    break;
                default:
                    break;
            }
            return ImportCSVToDatabase(dbTableName);
        }

        /// <summary> Private Method ImportCSVToDatabase
        /// Accepts destination DB table name and runs steps to select Import file, 
        /// process CSV, error check and import to DB and/or export errored entries.
        /// </summary>
        /// <param name="dbTableName">Database Table Name</param>
        /// <returns>bool of success/failure</returns>
        private static bool ImportCSVToDatabase(string dbTableName)
        {
            try
            {
                FileUtilities fileUtils = new FileUtilities();
                string fileName;
                DataTable dataTable;
                // open CSV
                fileName = fileUtils.OpenFile(null, "csv", "Select CSV");
                // Convert CSV to DataTable
                dataTable = CSVtoDataTable(fileName);
                // Update DataTable with correct values
                dataTable = UpdateDataTable(dataTable);
                //// Seperate errors to seperate file.
                //Tuple<DataTable, int> errorCount = SeperateErrors(dataTable, dbTableName);
                //// Check for duplicates in 'good' datatable
                //Tuple<int, int> rowCounts = DuplicatesCheck(dataTable, dbTableName);

                MessageBox.Show("New records imported to " + dbTableName + "." + Environment.NewLine
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);

                return true;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("Unable to retrieve import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Unable to import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Number + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Source + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }

        }

        /// <summary> Private Method CSVtoDataTable
        /// Takes csv filename and processes CSV into Datatable.
        /// Adds additional columns for DB with default values.
        /// </summary>
        /// <param name="fileName">Full path/filename</param>
        /// <returns>DataTable</returns>
        private static DataTable CSVtoDataTable(string fileName)
        {
            DataTable dataTable = new DataTable();
            using (StreamReader reader = new StreamReader(fileName))
            using (CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                // Do any configuration to `CsvReader` before creating CsvDataReader.
                using (CsvDataReader dataReader = new CsvDataReader(csv))
                {
                    dataTable.Columns.Add("UTR", typeof(int)); 
                    dataTable.Columns.Add("Project_Name", typeof(string));
                    dataTable.Columns.Add("Case_Reference", typeof(string));
                    dataTable.Columns.Add("Case_Status", typeof(string));
                    dataTable.Load(dataReader);
                }
            }
            //// Add actual columns needed for DB
            //DataColumn new1 = new System.Data.DataColumn("EmployerID", typeof(int));
            //DataColumn new2 = new System.Data.DataColumn("TaxYearID", typeof(int));
            //DataColumn new3 = new System.Data.DataColumn("NumberPaid", typeof(int));
            //DataColumn new4 = new System.Data.DataColumn("UpdatedBy", typeof(int));
            //DataColumn new5 = new System.Data.DataColumn("DateUpdated", typeof(DateTime));
            //DataColumn new6 = new System.Data.DataColumn("Duplicate", typeof(int));
            //DataColumn new7 = new System.Data.DataColumn("Value", typeof(int));
            //new1.DefaultValue = -1;
            //new2.DefaultValue = -1;
            //new3.DefaultValue = -1;
            //new4.DefaultValue = Global.PID;
            //new5.DefaultValue = DateTime.Now;
            //new6.DefaultValue = -1;
            //new7.DefaultValue = -1;
            //dataTable.Columns.Add(new1);
            //dataTable.Columns.Add(new2);
            //dataTable.Columns.Add(new3);
            //dataTable.Columns.Add(new4);
            //dataTable.Columns.Add(new5);
            //dataTable.Columns.Add(new6);
            //dataTable.Columns.Add(new7);

            return dataTable;
        }

        /// <summary> Private Method UpdateDataTable
        /// Gets lookup tables from DB and updates default values in datatable.
        /// </summary>
        /// <param name="dataTable">DataTable</param>
        /// <returns>DataTable</returns>
        private static DataTable UpdateDataTable(DataTable dataTable)
        {
            ////Lookups lu = new Lookups();
            //ImportFunctions lookups = new ImportFunctions();
            //lookups.GetInsertLookups();
            //DataTable employerLookup = lookups.ds.Tables[0];
            //DataTable taxyearLookup = lookups.ds.Tables[1];

            ////dataTable.PayeRef Lookup employerLookup.PAYERef Update dataTable.EmployerID to employerLookup.EmployerID
            //dataTable.AsEnumerable().Join(employerLookup.AsEnumerable(),    // data tables to be joined
            //        t1 => t1.Field<string>("PayeRef"),                      // fields to match from table 1
            //        t2 => t2.Field<string>("PAYERef"),                      // fields to match from table 1
            //        (t1, t2) => new { t1, t2 }).ToList()                    // magic bit
            //        .ForEach(o => o.t1["EmployerID"] = o.t2["EmployerID"]); // update field and value.

            ////dataTable.TaxYear Lookup taxyearLookup.TaxYearName Update dataTable.TaxYearID to taxyearLookup.TaxYearID
            //dataTable.AsEnumerable().Join(taxyearLookup.AsEnumerable(),
            //        t1 => t1.Field<string>("TaxYear"),
            //        t2 => t2.Field<string>("TaxYearName"),
            //        (t1, t2) => new { t1, t2 }).ToList()
            //        .ForEach(o => o.t1["TaxYearID"] = o.t2["TaxYearID"]);

            // dataTable.NINOsPaid to dataTable.NumberPaid
            dataTable.AsEnumerable().Join(dataTable.AsEnumerable(),
                    t1 => t1.Field<int>("UTR"),
                    t2 => t2.Field<int>("UTR"),
                    (t1, t2) => new { t1, t2 }).ToList()
                    .ForEach(o => o.t1["UTR"] = (int)o.t2["UTR"]);

            return dataTable;
        }

        /// <summary> Private Method SeperateErrors
        /// Duplicates updated DataTable for error checks.
        /// removes good data from errorChecked and exports as error file
        /// removes bad data from DataTable and returns.
        /// </summary>
        /// <param name="dataTable">Updated DataTable</param>
        /// <param name="dbTableName">DB Table Name for error Export</param>
        /// <returns>DataTable of good entries</returns>
        private static Tuple<DataTable, int> SeperateErrors(DataTable dataTable, string dbTableName)
        {
            Reports.DataGridReports reports = new Reports.DataGridReports();
            DataTable errorChecked = new DataTable();
            errorChecked = dataTable.Copy();
            // remove good data from errorChecked
            errorChecked.Rows.Cast<DataRow>().Where(x => x.Field<int>("UTR") > -1
                                                      && x.Field<int>("Caseflow_Reference") > -1
                                                      && x.Field<int>("Project_Name") > -1).ToList().ForEach(x => x.Delete());
            errorChecked.AcceptChanges();
            // send errorChecked to CSV
            int errorCount = errorChecked.Rows.Count;
            if (errorCount > 0) { ExportImportErrors(dbTableName, errorChecked, errorCount); }

            // remove bad data from dataTable
            dataTable.Rows.Cast<DataRow>().Where(x => x.Field<int>("UTR") < 0
                                                   || x.Field<int>("Caseflow_Reference") < 0
                                                   || x.Field<int>("Project_Name") < 0).ToList().ForEach(x => x.Delete());
            dataTable.AcceptChanges();

            Tuple<DataTable, int> rowCounts = new Tuple<DataTable, int>(dataTable, errorCount);
            return rowCounts;
        }

        /// <summary> Private Method ExportImportErrors
        /// Takes errored DataTable and exports to error file. 
        /// <param name="dbTableName">Destination DB Table Name</param>
        /// <param name="errorChecked">DataTable of errored entries</param>
        /// <returns>Bool of success/failure</returns>
        private static void ExportImportErrors(string dbTableName, DataTable errorChecked, int errorCount)
        {
            MessageBox.Show("There are " + errorCount + " PAYE refs not found." + Environment.NewLine + Environment.NewLine +
                "Please select a location to export records." + Environment.NewLine
                , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);

            FileUtilities fileUtils = new FileUtilities();
            //string folderName = @"D:\Users\U.6074887\Outputs\";
            string folderName = fileUtils.PickFolder(null);
            ReportFunctions reports = new ReportFunctions();
            reports.DataTableToCSV(folderName, dbTableName, errorChecked);
        }

        /// <summary> Private Method DuplicatesCheck
        /// Checks good table for existing duplicates in dbTableName
        /// </summary>
        /// <param name="dataTable">Updated DataTable</param>
        /// <param name="dbTableName">DB Table Name for error Export</param>
        /// <returns>Tuple of rowcounts</returns>
        private static Tuple<int, int> DuplicatesCheck(DataTable dataTable, string dbTableName)
        {
            ImportFunctions lookups = new ImportFunctions();
            lookups.GetDuplicateChecks(dbTableName);
            DataTable existingData = lookups.ds.Tables[0];

            // Match ID's if both match mark as a duplicate.
            dataTable.AsEnumerable().Join(existingData.AsEnumerable(),                              // data tables to be joined
                    t1 => new { x = t1.Field<int>("UTR"), y = t1.Field<int>("Caseflow_Reference") },  // fields to match from table 1
                    t2 => new { x = t2.Field<int>("UTR"), y = t2.Field<int>("Caseflow_Reference") },  // fields to match from table 2
                    (t1, t2) => new { t1, t2 }).ToList()                                            // magic bit
                    .ForEach(o => o.t1["Duplicate"] = 1);                                           // update field and value.
            // Match ID's and add Numbers Paid value for next check.
            dataTable.AsEnumerable().Join(existingData.AsEnumerable(),                              // data tables to be joined
                    t1 => new { x = t1.Field<int>("UTR"), y = t1.Field<int>("Caseflow_Reference") },  // fields to match from table 1
                    t2 => new { x = t2.Field<int>("UTR"), y = t2.Field<int>("Caseflow_Reference") },  // fields to match from table 2
                    (t1, t2) => new { t1, t2 }).ToList()                                            // magic bit
                    .ForEach(o => o.t1["Project_Name"] = o.t2["Project_Name"]);                              // update field and value.

            DataTable duplicates = new DataTable();
            duplicates = dataTable.Copy();

            // remove duplicates from dataTable and import new entries to DB
            dataTable.Rows.Cast<DataRow>().Where(x => x.Field<int>("Duplicate") == 1).ToList().ForEach(x => x.Delete());
            dataTable.AcceptChanges();
            DataTableToSQL(dataTable, dbTableName);

            // Remove new data from Duplicates
            duplicates.Rows.Cast<DataRow>().Where(x => x.Field<int>("Duplicate") == -1).ToList().ForEach(x => x.Delete());
            duplicates.AcceptChanges();

            // Remove date where NumberPaid matches existing value: No need to update
            duplicates.Rows.Cast<DataRow>().Where(x => x.Field<int>("Project_Name") == x.Field<int>("Project_Name")).ToList().ForEach(x => x.Delete());
            duplicates.AcceptChanges();

            // Remove columns not needed by TableValuedParameter
            duplicates.Columns.Remove("UTR");
            duplicates.Columns.Remove("Caseflow_Reference");
            duplicates.Columns.Remove("Project_Name");
            //duplicates.Columns.Remove("EmployerID");
            //duplicates.Columns.Remove("TaxYearID");
            //duplicates.Columns.Remove("NumberPaid");

            //duplicates.Columns.Remove("UpdatedBy");
            //duplicates.Columns.Remove("DateUpdated");
            duplicates.Columns.Remove("Duplicate");
            duplicates.Columns.Remove("Case_Status");

            // send remaining updates to seperate Merge process.
            lookups.UpdateToSQL(duplicates, dbTableName);

            Tuple<int, int> rowCounts = new Tuple<int, int>(dataTable.Rows.Count, duplicates.Rows.Count);
            return rowCounts;
        }

        /// <summary> Private Method DataTableToSQL
        /// Takes Good DataTable and imports to DB table. 
        /// </summary>
        /// <param name="dataTable">DataTable of Good entries</param>
        /// <param name="dbTableName">Destination DB Table Name</param>
        /// <returns>Bool of success/failure</returns>
        private static bool DataTableToSQL(DataTable dataTable, string dbTableName)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        //create object of SqlBulkCopy which help to insert  
                        SqlBulkCopy objbulk = new SqlBulkCopy(con);
                        //assign Destination table name  
                        objbulk.DestinationTableName = "[dbo]." + dbTableName;
                        objbulk.ColumnMappings.Add("UTR", "[UTR]");
                        objbulk.ColumnMappings.Add("Caseflow_Reference", "[Caseflow_Reference]");
                        objbulk.ColumnMappings.Add("Project_Name", "[Project_Name]");
                        objbulk.ColumnMappings.Add("Case_Status", "[Case_Status]");
                        con.Open();
                        //insert bulk Records into DataBase.  
                        objbulk.WriteToServer(dataTable);
                        con.Close();
                        return true;
                    }
                }
            }
            catch (SqlException e)
            {
                System.Windows.MessageBox.Show("Unable to insert data to database." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
        }

        /// <summary> Private Method GetInsertLookups
        /// Gets lookup tables from DB for updating DataTable defauklt entries. 
        /// </summary>
        /// <returns></returns>
        private void GetInsertLookups()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        //cmd.CommandText = "uspGetMainWindowCombos";
                        cmd.CommandText = "[ExPat].[uspGetInsertLookups]";
                        //SqlParameter prm = cmd.Parameters.Add("@inListType", SqlDbType.Text);
                        //prm.Value = "ALL";
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;
                        con.Open();
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(ds);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve Lookups for Insert process." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
            }

        }

        /// <summary> Private Method GetDuplicateChecks
        /// Gets existing ID's in tableName to check for duplicates. 
        /// </summary>
        /// <param name="dbTableName">Destination DB Table Name</param>
        /// <returns></returns>
        private void GetDuplicateChecks(string tableName)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        //cmd.CommandText = "uspGetMainWindowCombos";
                        cmd.CommandText = "[ExPat].[uspCheckForDuplicates]";
                        SqlParameter prm = cmd.Parameters.Add("@inTableName", SqlDbType.Text);
                        prm.Value = "[ExPat]." + tableName;
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;
                        con.Open();
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(ds);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve Lookups for Insert process." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
            }

        }

        /// <summary> Private Method GetDuplicateChecks
        /// Gets existing ID's in tableName to check for duplicates. 
        /// </summary>
        /// <param name="dbTableName">Destination DB Table Name</param>
        /// <returns></returns>
        private void UpdateToSQL(DataTable dataTable, string tableName)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "[ExPat].[uspUpdate" + tableName.Substring(tableName.Length - (tableName.Length - 1));
                        SqlParameter tvp = cmd.Parameters.AddWithValue("@DataTable", dataTable);
                        tvp.SqlDbType = SqlDbType.Structured;
                        tvp.TypeName = "[ExPat].[uttAnnualImports]";
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to update existing entries." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
            }

        }
    }
}

