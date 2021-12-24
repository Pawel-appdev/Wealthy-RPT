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
    class ImportCRH : Reports.ReportFunctions
    {
        SqlDataAdapter adapter = new SqlDataAdapter();
        public DataSet ds = new DataSet();

        /// <summary> Public Method ImportEmployerTest
        /// Allows import of Employer data and outputs a DB ready file for error checking.
        /// </summary>
        /// <returns>bool of success/failure</returns>
        public static bool ImportEmployerTest()
        {
            try
            {
                FileUtilities fileUtils = new FileUtilities();
                string fileName;
                DataTable dataTable;
                // open CSV
                fileName = fileUtils.OpenFile(null, "csv", "Select Caseflow History CSV");
                // Convert CSV to DataTable
                dataTable = CSVtoDataTable(fileName);
                ExportEmployers(dataTable, "Employer Import Check Pre Conversion");
                // Convert values for DB
                dataTable = DataTableToDBReady(dataTable);
                // Output to CSV for checking
                ExportEmployers(dataTable, "Employer Import Check Post Conversion");

                MessageBox.Show(dataTable.Rows.Count + " records exported for checking."
                        + Environment.NewLine + Environment.NewLine
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
            catch (Exception ex)
            {
                MessageBox.Show("Unable to import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Source + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
        }


        /// <summary> Public Method ImportMasterEmployerFile
        /// Accepts destination DB table name and runs steps to select Import file, 
        /// process CSV, error check and import to DB and/or export errored entries.
        /// </summary>
        /// <returns>bool of success/failure</returns>
        public static bool ImportCRMMRiskHistory()
        {
            try
            {
                FileUtilities fileUtils = new FileUtilities();
                string fileName;
                DataTable dataTable;
                // open CSV
                fileName = fileUtils.OpenFile(null, "csv", "Select CRMM Risk History CSV");
                // Convert CSV to DataTable
                dataTable = CSVtoDataTable(fileName);
                // delete previous data if needed
                var selectedOption = MessageBox.Show("Do you wish to delete the existing CRMM Risk History Data.", "Weathy RPT", MessageBoxButton.YesNo, MessageBoxImage.Question);
                switch (selectedOption)
                {
                    case MessageBoxResult.Yes:
                        DeletePrevCHData();
                        break;
                    case MessageBoxResult.No:
                        break;
                }
                //// Convert values for DB
                //dataTable = DataTableToDBReady(dataTable);
                // DataTable to SQL USP
                DataTableToSQL(dataTable);

                MessageBox.Show(dataTable.Rows.Count + " records imported to Employers."
                        + Environment.NewLine + Environment.NewLine
                        , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);

                return true;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("Unable to retrieve CRMM Risk History import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Unable to import CRMM Risk History data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Number + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to import CRMM Risk" +
                    " History data." + Environment.NewLine + Environment.NewLine
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
                    dataTable.Columns.Add("RiskID", typeof(string));
                    dataTable.Columns.Add("RiskStatus", typeof(string));
                    dataTable.Columns.Add("SettledDate", typeof(DateTime));

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
                    objbulk.DestinationTableName = "[dbo].[tblCRMMRisks]";

                    objbulk.ColumnMappings.Add("UTR", "[UTR]");
                    objbulk.ColumnMappings.Add("RiskID", "[RiskID]");
                    objbulk.ColumnMappings.Add("RiskStatus", "[RiskStatus]");
                    objbulk.ColumnMappings.Add("SettledDate", "[SettledDate]");

                    con.Open();
                    //insert bulk Records into DataBase.  
                    objbulk.WriteToServer(dataTable);
                    con.Close();
                    return true;
                }
            }
        }

        private static bool DeletePrevCHData()
        {
            SqlConnection con = new SqlConnection(Global.ConnectionString);
            try
            {
                SqlCommand cmd = new SqlCommand("qryDeletePreviousCRMMRiskHistory", con);
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
