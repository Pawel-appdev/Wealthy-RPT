using CsvHelper;
using ExPat_UI.Reports;
using ExPat_UI.Utilities;
using Wealthy_RPT;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Windows;

namespace ExPat_UI.Imports
{
    class ImportUH : Reports.ReportFunctions
    {
        SqlDataAdapter adapter = new SqlDataAdapter();
        public DataSet ds = new DataSet();

        /// <summary> Public Method ImportMasterEmployerFile
        /// Accepts destination DB table name and runs steps to select Import file, 
        /// process CSV, error check and import to DB and/or export errored entries.
        /// </summary>
        /// <returns>bool of success/failure</returns>
        public static bool ImportUpstreamHistory()
        {
            try
            {
                FileUtilities fileUtils = new FileUtilities();
                string fileName;
                DataTable dataTable;
                // open CSV
                fileName = fileUtils.OpenFile(null, "csv", "Select Upstream History CSV");
                // Convert CSV to DataTable
                dataTable = CSVtoDataTable(fileName);
                // delete previous data if needed
                var selectedOption = MessageBox.Show("Do you wish to delete the existing Upstream History Data.", "Weathy RPT", MessageBoxButton.YesNo, MessageBoxImage.Question);
                switch (selectedOption)
                {
                    case MessageBoxResult.Yes:
                        DeletePrevUHData();
                        break;
                    case MessageBoxResult.No:
                        break;
                }
                //// Convert values for DB
                //dataTable = DataTableToDBReady(dataTable);
                // DataTable to SQL USP
                DataTableToSQL(dataTable);

                MessageBox.Show(dataTable.Rows.Count + " records imported to Upstream History."
                        + Environment.NewLine + Environment.NewLine
                        , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);

                return true;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("Unable to retrieve Upstream History import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Unable to import Upstream History data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Number + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to import Upstream History data." + Environment.NewLine + Environment.NewLine
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
                    dataTable.Columns.Add("OTM_Reference", typeof(string));
                    dataTable.Columns.Add("Upstream_Project_Name", typeof(string));
                    dataTable.Columns.Add("Case_Status", typeof(string));
                    //dataTable.Columns.Add("EntityID", typeof(int));
                    //dataTable.Columns.Add("CountryID", typeof(int));
                    //dataTable.Columns.Add("COTAXUTR", typeof(string));
                    //dataTable.Columns.Add("CustomerGroupID", typeof(int));
                    //dataTable.Columns.Add("IndustrySectorID", typeof(int));
                    //// Nullable columns read as strings convert later
                    //dataTable.Columns.Add("DedicatedScheme", typeof(string));
                    //dataTable.Columns.Add("Appendix4", typeof(string));
                    //dataTable.Columns.Add("Appendix5", typeof(string));
                    //dataTable.Columns.Add("Appendix6", typeof(string));
                    //dataTable.Columns.Add("Appendix7A", typeof(string));
                    //dataTable.Columns.Add("Appendix7B", typeof(string));
                    //dataTable.Columns.Add("Appendix8", typeof(string));
                    //dataTable.Columns.Add("NonResidentDirector", typeof(string));
                    //dataTable.Columns.Add("ConcessionaryCashBasis", typeof(string));
                    //dataTable.Columns.Add("BespokeScaleRateAgreement", typeof(string));
                    //dataTable.Columns.Add("S690", typeof(string));
                    //dataTable.Columns.Add("Notes", typeof(string));
                    //dataTable.Columns.Add("ActiveScheme", typeof(string));

                    dataTable.Load(dataReader);
                }
            }
            
            return dataTable;
        }

        /// <summary> DataTableToDBReady
        /// Work through each row and update nullable fields with correct value. 
        /// </summary>
        /// <param name="dataTable">DataTable to be edited</param>
        /// <returns></returns>
        private static DataTable DataTableToDBReady(DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                BitToDBNullableBit(row, "DedicatedScheme");
                DateToDBNullableDate(row, "Appendix4");
                DateToDBNullableDate(row, "Appendix5");
                DateToDBNullableDate(row, "Appendix6");
                DateToDBNullableDate(row, "Appendix7a");
                DateToDBNullableDate(row, "Appendix7b");
                DateToDBNullableDate(row, "Appendix8");
                DateToDBNullableDate(row, "NonResidentDirector");
                DateToDBNullableDate(row, "ConcessionaryCashBasis");
                DateToDBNullableDate(row, "BespokeScaleRateAgreement");
                DateToDBNullableDate(row, "S690");
                BitToDBNullableBit(row, "ActiveScheme");
            }

            return dataTable;
        }

        /// <summary> DateToDBNullableDate
        /// Takes columnName and data row converts String of Date to either DateTime or DBNull
        /// </summary>
        /// <param name="row">DataRow from DataTable</param>
        /// <param name="columnName">ColumnName to be edited</param>
        /// <returns></returns>
        private static void DateToDBNullableDate(DataRow row, string columnName)
        {
            DateTime nullDate;
            bool testAppendix4 = DateTime.TryParse(row[columnName].ToString(), out nullDate);
            if (testAppendix4)
            {
                row["Temp" + columnName] = DateTime.Parse(row[columnName].ToString());
            }
            else
            {
                row["Temp" + columnName] = DBNull.Value;
            }
        }

        /// <summary> BitToDBNullableBit
        /// Takes columnName and data row converts String of Bit to either Bit or DBNull
        /// </summary>
        /// <param name="row">DataRow from DataTable</param>
        /// <param name="columnName">ColumnName to be edited</param>
        /// <returns></returns>
        private static void BitToDBNullableBit(DataRow row, string columnName)
        {
            switch (row[columnName].ToString())
            {
                case "1":
                    row["Temp" + columnName] = 1;
                    break;
                case "0":
                    row["Temp" + columnName] = 0;
                    break;
                default:
                    row["Temp" + columnName] = DBNull.Value;
                    break;
            }
        }

        /// <summary> Private Method ExportImportErrors
        /// Takes errored DataTable and exports to error file. 
        /// <param name="dbTableName">Destination DB Table Name</param>
        /// <param name="errorChecked">DataTable of errored entries</param>
        /// <returns>Bool of success/failure</returns>
        private static void ExportEmployers(DataTable errorChecked, string fileName)
        {
            FileUtilities fileUtils = new FileUtilities();
            string folderName = fileUtils.PickFolder(null);
            ReportFunctions reports = new ReportFunctions();
            reports.DataTableToCSV(folderName, fileName, errorChecked);
        }


        /// <summary> Private Method DataTableToSQL
        /// Takes Employers DataTable and imports to DB table. 
        /// </summary>
        /// <param name="dataTable">DataTable for import</param>
        /// <returns>Bool of success/failure</returns>
        private static bool DataTableToSQL(DataTable dataTable)
        {

            using (SqlConnection con = new SqlConnection(Global.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    //create object of SqlBulkCopy which help to insert  
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    //assign Destination table name  
                    objbulk.DestinationTableName = "[dbo].[tblUpstreamHistory]";

                    objbulk.ColumnMappings.Add("UTR", "[UTR]");
                    objbulk.ColumnMappings.Add("OTM_Reference", "[OTM_Reference]");
                    objbulk.ColumnMappings.Add("Upstream_Project_Name", "[Upstream_Project_Name]");
                    objbulk.ColumnMappings.Add("Case_Status", "[Case_Status]");

                    con.Open();
                    //insert bulk Records into DataBase.  
                    objbulk.WriteToServer(dataTable);
                    con.Close();
                    return true;
                }
            }
        }

        private static bool DeletePrevUHData()
        {
            SqlConnection con = new SqlConnection(Global.ConnectionString);
            try
            {
                SqlCommand cmd = new SqlCommand("qryDeletePreviousUpstreamHistory", con);
                cmd.CommandTimeout = Global.TimeOut;
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                Console.Write("SQL error: " + ex.Message);
                con.Close();
                return false;
            }
            catch (Exception ex)
            {
                Console.Write("Error: " + ex.Message);
                con.Close();
                return false;
            }

            return true;
        }
    }
}
