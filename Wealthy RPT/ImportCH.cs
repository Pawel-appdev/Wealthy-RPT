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
    class ImportCH : Reports.ReportFunctions
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
        public static bool ImportCaseflowHistory()
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
                // delete previous data if needed
                var selectedOption = MessageBox.Show("Do you wish to delete the existing Caseflow History Data.", "Weathy RPT", MessageBoxButton.YesNo, MessageBoxImage.Question);
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
                MessageBox.Show("Unable to retrieve Caseflow History import data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Unable to import Caseflow History data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Number + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to import Caseflow History data." + Environment.NewLine + Environment.NewLine
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
                    dataTable.Columns.Add("Caseflow_Reference", typeof(string));
                    dataTable.Columns.Add("Project_Name", typeof(string));
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
            //// Add actual columns needed for conversion to DB values allowing nulls
            //DataColumn DedicatedScheme = new System.Data.DataColumn("TempDedicatedScheme", typeof(int));
            //DataColumn Appendix4 = new System.Data.DataColumn("TempAppendix4", typeof(DateTime));
            //DataColumn Appendix5 = new System.Data.DataColumn("TempAppendix5", typeof(DateTime));
            //DataColumn Appendix6 = new System.Data.DataColumn("TempAppendix6", typeof(DateTime));
            //DataColumn Appendix7a = new System.Data.DataColumn("TempAppendix7a", typeof(DateTime));
            //DataColumn Appendix7b = new System.Data.DataColumn("TempAppendix7b", typeof(DateTime));
            //DataColumn Appendix8 = new System.Data.DataColumn("TempAppendix8", typeof(DateTime));
            //DataColumn NonResidentDirector = new System.Data.DataColumn("TempNonResidentDirector", typeof(DateTime));
            //DataColumn ConcessionaryCashBasis = new System.Data.DataColumn("TempConcessionaryCashBasis", typeof(DateTime));
            //DataColumn BespokeScaleRateAgreement = new System.Data.DataColumn("TempBespokeScaleRateAgreement", typeof(DateTime));
            //DataColumn S690 = new System.Data.DataColumn("TempS690", typeof(DateTime));
            //DataColumn ActiveScheme = new System.Data.DataColumn("TempActiveScheme", typeof(int));

            //DedicatedScheme.AllowDBNull = true;
            //Appendix4.AllowDBNull = true;
            //Appendix5.AllowDBNull = true;
            //Appendix6.AllowDBNull = true;
            //Appendix7a.AllowDBNull = true;
            //Appendix7b.AllowDBNull = true;
            //Appendix8.AllowDBNull = true;
            //NonResidentDirector.AllowDBNull = true;
            //ConcessionaryCashBasis.AllowDBNull = true;
            //BespokeScaleRateAgreement.AllowDBNull = true;
            //S690.AllowDBNull = true;
            //ActiveScheme.AllowDBNull = true;

            //dataTable.Columns.Add(DedicatedScheme);
            //dataTable.Columns.Add(Appendix4);
            //dataTable.Columns.Add(Appendix5);
            //dataTable.Columns.Add(Appendix6);
            //dataTable.Columns.Add(Appendix7a);
            //dataTable.Columns.Add(Appendix7b);
            //dataTable.Columns.Add(Appendix8);
            //dataTable.Columns.Add(NonResidentDirector);
            //dataTable.Columns.Add(ConcessionaryCashBasis);
            //dataTable.Columns.Add(BespokeScaleRateAgreement);
            //dataTable.Columns.Add(S690);
            //dataTable.Columns.Add(ActiveScheme);

            //// system generated DB columns
            //DataColumn UpdatedBy = new System.Data.DataColumn("UpdatedBy", typeof(int));
            //DataColumn DateUpdated = new System.Data.DataColumn("DateUpdated", typeof(DateTime));
            //UpdatedBy.DefaultValue = Global.PID;
            //DateUpdated.DefaultValue = DateTime.Now;
            //dataTable.Columns.Add(UpdatedBy);
            //dataTable.Columns.Add(DateUpdated);

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
                    objbulk.DestinationTableName = "[dbo].[tblCaseflowHistory]";

                    //   [ExPat].[Employers]
                    //   [EmployerID][int] IDENTITY(-1, 1) NOT NULL,
                    //   [PAYERef] [varchar](50) NOT NULL,
                    //   [EmployerName] [varchar](50) NOT NULL,
                    //   [ParentGroupID] [int] NOT NULL,
                    //   [IsParent] [bit] NOT NULL,

                    objbulk.ColumnMappings.Add("UTR", "[UTR]");
                    objbulk.ColumnMappings.Add("Caseflow_Reference", "[Caseflow_Reference]");
                    objbulk.ColumnMappings.Add("Project_Name", "[Project_Name]");
                    objbulk.ColumnMappings.Add("Case_Status", "[Case_Status]");

                    ////   [EntityID] [int] NOT NULL,
                    ////   [CountryID] [int] NOT NULL,
                    ////   [COTAXUTR] [varchar](50) NOT NULL,
                    ////   [CustomerGroupID] [int] NOT NULL,
                    ////   [IndustrySectorID] [int] NOT NULL,

                    //objbulk.ColumnMappings.Add("EntityID", "[EntityID]");
                    //objbulk.ColumnMappings.Add("CountryID", "[CountryID]");
                    //objbulk.ColumnMappings.Add("COTAXUTR", "[COTAXUTR]");
                    //objbulk.ColumnMappings.Add("CustomerGroupID", "[CustomerGroupID]");
                    //objbulk.ColumnMappings.Add("IndustrySectorID", "[IndustrySectorID]");

                    ////   [DedicatedScheme] [bit] NULL,
                    ////   [Appendix4] [date] NULL,
                    ////   [Appendix5] [date] NULL,
                    ////   [Appendix6] [date] NULL,
                    ////   [Appendix7A] [date] NULL,
                    ////   [Appendix7B] [date] NULL,
                    ////   [Appendix8] [date] NULL,
                    ////   [NonResidentDirector] [bit] NULL,
                    ////   [ConcessionaryCashBasis] [date] NULL,
                    ////   [BespokeScaleRateAgreement] [date] NULL,
                    ////   [S690] [date] NULL,
                    //objbulk.ColumnMappings.Add("TempDedicatedScheme", "[DedicatedScheme]");
                    //objbulk.ColumnMappings.Add("TempAppendix4", "[Appendix4]");
                    //objbulk.ColumnMappings.Add("TempAppendix5", "[Appendix5]");
                    //objbulk.ColumnMappings.Add("TempAppendix6", "[Appendix6]");
                    //objbulk.ColumnMappings.Add("TempAppendix7A", "[Appendix7A]");
                    //objbulk.ColumnMappings.Add("TempAppendix7B", "[Appendix7B]");
                    //objbulk.ColumnMappings.Add("TempAppendix8", "[Appendix8]");
                    //objbulk.ColumnMappings.Add("TempNonResidentDirector", "[NonResidentDirector]");
                    //objbulk.ColumnMappings.Add("TempConcessionaryCashBasis", "[ConcessionaryCashBasis]");
                    //objbulk.ColumnMappings.Add("TempBespokeScaleRateAgreement", "[BespokeScaleRateAgreement]");
                    //objbulk.ColumnMappings.Add("TempS690", "[S690]");
                    ////   [Notes] [varchar](50) NULL,
                    ////   [ActiveScheme] [bit] NULL,
                    ////   [UpdatedBy] [int] NOT NULL,
                    ////   [DateUpdated] [date] NOT NULL,
                    //objbulk.ColumnMappings.Add("Notes", "[Notes]");
                    //objbulk.ColumnMappings.Add("TempActiveScheme", "[ActiveScheme]");
                    //objbulk.ColumnMappings.Add("UpdatedBy", "[UpdatedBy]");
                    //objbulk.ColumnMappings.Add("DateUpdated", "[DateUpdated]");

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
                SqlCommand cmd = new SqlCommand("qryDeletePreviousCaseflowHistory", con);
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
