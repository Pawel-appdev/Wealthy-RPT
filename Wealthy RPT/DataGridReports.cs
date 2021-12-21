//using ExPat_UI.Model;
using ExPat_UI.Utilities;
using Wealthy_RPT;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ExPat_UI.Reports
{
    class DataGridReports : ReportFunctions
    {
        SqlDataAdapter adapter = new SqlDataAdapter();
        public DataSet ds = new DataSet();

        /// <summary> dgEmployersToCSV
        /// Takes DataGridView from main Window and exports it to CSV 
        /// </summary>
        /// <param name="dgEmployers">DataGrid from main window</param>
        /// <returns>Bool of success/failure</returns>
        public static bool dgEmployersToCSV(DataGrid dgEmployers)
        {
            try
            {
                FileUtilities fileUtils = new FileUtilities();
                string folderName = fileUtils.PickFolder(null);

                ReportFunctions reportFunctions = new ReportFunctions();
                DataTable dtEmployers = reportFunctions.DataGridtoDataTable(dgEmployers);
                reportFunctions.DataTableToCSV(folderName, "ExPat", dtEmployers);

                MessageBox.Show(dtEmployers.Rows.Count + " records exported to " + folderName + "."
                    + Environment.NewLine + Environment.NewLine
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Source + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
        }

        /// <summary> dgEmployersToCSV
        /// Takes DataGridView from main Window and exports it to CSV 
        /// </summary>
        /// <param name="filterValues">Employer Object of search/filter values</param>
        /// <returns>Bool of success/failure</returns>
        /// 
        //Removed by Liam
        //public static bool ReportEmployersViewToCSV(Employers.Employer filterValues)
        //{
        //    try
        //    {
        //        FileUtilities fileUtils = new FileUtilities();
        //        string folderName = fileUtils.PickFolder(null);

        //        DataGridReports data = new DataGridReports();
        //        data.GetViewData(filterValues);
        //        DataTable dtEmployers = data.ds.Tables[0];

        //        ReportFunctions reportFunctions = new ReportFunctions();
        //        reportFunctions.DataTableToCSV(folderName, "ExPat Report", dtEmployers);

        //        MessageBox.Show(dtEmployers.Rows.Count + " records exported to " + folderName + "."
        //            + Environment.NewLine + Environment.NewLine
        //            , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
        //        return true;
        //    }
        //    catch (FileNotFoundException ex)
        //    {
        //        MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
        //            + "Error: " + ex.Message
        //            , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
        //        return false;
        //    }
        //    catch (SqlException ex)
        //    {
        //        MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
        //            + "Error: " + ex.Number + Environment.NewLine + ex.Message
        //            , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
        //        return false;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
        //            + "Error: " + ex.Source + Environment.NewLine + ex.Message
        //            , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
        //        return false;
        //    }
        //}

        /// <summary> ReportAdditionalDataToCSV
        /// Runs Procedure to gather data from additional Tables for last 4 years and exports it to CSV 
        /// </summary>
        /// <returns>Bool of success/failure</returns>
        public static bool ReportAdditionalDataToCSV()
        {
            try
            {
                FileUtilities fileUtils = new FileUtilities();
                string folderName = fileUtils.PickFolder(null);

                DataGridReports data = new DataGridReports();
                data.GetAdditionalTableData();
                DataTable dtEmployers = data.ds.Tables[0];

                ReportFunctions reportFunctions = new ReportFunctions();
                reportFunctions.DataTableToCSV(folderName, "ExPat Last 4 years", dtEmployers);

                MessageBox.Show(dtEmployers.Rows.Count + " records exported to " + folderName + "."
                    + Environment.NewLine + Environment.NewLine
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Number + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to retrieve report data." + Environment.NewLine + Environment.NewLine
                    + "Error: " + ex.Source + Environment.NewLine + ex.Message
                    , Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Stop);
                return false;
            }
        }

        /// <summary> GetViewData
        /// SQL procedure to retrieve data based on filters in view
        /// </summary>
        /// <param name="filterValues">Employer Object of search/filter values</param>
        /// <returns></returns>
        /// 
        //Removed by Liam
        //private void GetViewData(Employers.Employer filterValues)
        //{
        //    using (SqlConnection con = new SqlConnection(Global.ConnectionString))
        //    {
        //        using (SqlCommand cmd = new SqlCommand())
        //        {
        //            cmd.CommandText = "[ExPat].[uspReportView]";

        //            // Search Inputs
        //            cmd.Parameters.Add("@inSearchString", SqlDbType.Text).Value = filterValues.SearchString;
        //            cmd.Parameters.Add("@inSearchColumn", SqlDbType.Text).Value = filterValues.SearchColumn;
        //            // Filter inputs
        //            cmd.Parameters.Add("@inEntity", SqlDbType.Int).Value = filterValues.EntityID;
        //            cmd.Parameters.Add("@inCustomerGroup", SqlDbType.Int).Value = filterValues.CustomerGroupID;
        //            cmd.Parameters.Add("@inIndustrySector", SqlDbType.Int).Value = filterValues.IndustrySectorID;
        //            cmd.Parameters.Add("@inAppendix4", SqlDbType.Bit).Value = filterValues.Appendix4Flag ? 1 : 0;
        //            cmd.Parameters.Add("@inAppendix5", SqlDbType.Bit).Value = filterValues.Appendix5Flag ? 1 : 0;
        //            cmd.Parameters.Add("@inAppendix6", SqlDbType.Bit).Value = filterValues.Appendix6Flag ? 1 : 0;
        //            cmd.Parameters.Add("@inAppendix7A", SqlDbType.Bit).Value = filterValues.Appendix7AFlag ? 1 : 0;
        //            cmd.Parameters.Add("@inAppendix7B", SqlDbType.Bit).Value = filterValues.Appendix7BFlag ? 1 : 0;
        //            cmd.Parameters.Add("@inAppendix8", SqlDbType.Bit).Value = filterValues.Appendix8Flag ? 1 : 0;
        //            cmd.Parameters.Add("@inDedicatedScheme", SqlDbType.Bit).Value = filterValues.DedicatedScheme ? 1 : 0;
        //            cmd.Parameters.Add("@inNonResidentDirector", SqlDbType.Bit).Value = filterValues.NonResidentDirector ? 1 : 0;
        //            cmd.Parameters.Add("@inConcessionaryCashBasis", SqlDbType.Bit).Value = filterValues.ConcessionaryCashBasisFlag ? 1 : 0;
        //            cmd.Parameters.Add("@inBespokeScaleRateAgreement", SqlDbType.Bit).Value = filterValues.BespokeScaleRateAgreementFlag ? 1 : 0;
        //            cmd.Parameters.Add("@inS690", SqlDbType.Bit).Value = filterValues.S690Flag ? 1 : 0;
        //            cmd.Parameters.Add("@inS828a", SqlDbType.Bit).Value = filterValues.S828AFlag ? 1 : 0;
        //            cmd.Parameters.Add("@inActive", SqlDbType.Bit).Value = filterValues.ActiveScheme ? 1 : 0;
        //            cmd.Parameters.Add("@inInactiveDate", SqlDbType.Bit).Value = filterValues.IsInactive ? 1 : 0;

        //            cmd.Connection = con;
        //            cmd.CommandType = CommandType.StoredProcedure;
        //            con.Open();
        //            adapter = new SqlDataAdapter(cmd);
        //            adapter.Fill(ds);
        //            con.Close();
        //        }
        //    }
        //}

        /// <summary> GetAdditionalTableData
        /// SQL procedure to retrieve data from additional tables for last 4 years
        /// </summary>
        /// <returns></returns>
        private void GetAdditionalTableData()
        {
            using (SqlConnection con = new SqlConnection(Global.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "[ExPat].[uspReportAdditionalDataTables]";
                    cmd.Connection = con;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(ds);
                    con.Close();
                }
            }
        }

    }
}
