using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for PeriodImport.xaml
    /// </summary>
    public partial class PeriodImport : Window
    {
        public PeriodImport()
        {
            InitializeComponent();

            PopulateCombos();
        }

        public bool importdatafromexcel(string strExcelFilePath, string strExcelDataQuery, string strTable, string dataname)
        {
            // make sure your sheet name is correct, here sheet name is sheet1, so you can change your sheet name if have different
            // string myexceldataquery = "select student,rollno,course from [sheet1$]";
            try
            {
                //Import from Excel

                //set variables
                string TableName = strTable;
                String SchemaName = @"dbo";
                string fileFullPath = strExcelFilePath;

                //first clear table
                string sclearsql = "delete from " + strTable;
                SqlConnection sqlconn = new SqlConnection(Global.ConnectionString);
                SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
                sqlconn.Open();
                sqlcmd.ExecuteNonQuery();
                sqlconn.Close();

                //Create Connection to SQL Server Database 
                SqlConnection SQLConnection = new SqlConnection(Global.ConnectionString);

                //Create Excel Connection
                string ConStr;
                string HDR;
                HDR = "YES";
                ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + fileFullPath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
                OleDbConnection cnn = new OleDbConnection(ConStr);


                //Get data from Excel Sheet to DataTable
                OleDbConnection Conn = new OleDbConnection(ConStr);
                Conn.Open();
                OleDbCommand oconn = new OleDbCommand(strExcelDataQuery, Conn);
                OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                DataTable dt = new DataTable();
                adp.Fill(dt);
                Conn.Close();

                SQLConnection.Open();
                //Load Data from DataTable to SQL Server Table.
                using (SqlBulkCopy BC = new SqlBulkCopy(SQLConnection))
                {
                    BC.DestinationTableName = SchemaName + "." + TableName;
                    foreach (var column in dt.Columns)
                        BC.ColumnMappings.Add(column.ToString(), column.ToString());
                    BC.WriteToServer(dt);
                }
                SQLConnection.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }

        }

        private string PickFile(String DataRequired)
        {
            string info;
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Title = DataRequired;
            ofd.InitialDirectory = Global.DefaultImportFolder.ToString();
            ofd.FileName = "";
            ofd.Filter = "Excel Files|*.xls;";

            // Show open file dialog box
            Nullable<bool> result = ofd.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                info = "" + ofd.FileName + "";
            }
            else
            {
                info = "";
            }
            return info;
        }

        private void btnRisks_Click(object sender, RoutedEventArgs e)
        {
            this.txtRisks.Text = PickFile("Select Risks file");
        }

        private void btnAvoidanceScores_Click(object sender, RoutedEventArgs e)
        {
            this.txtAvoidanceScores.Text = PickFile("Select Avoidance Scores file");
        }

        private void btnCRMMData_Click(object sender, RoutedEventArgs e)
        {
            this.txtCRMMData.Text = PickFile("Select CRMM file");
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            string strYear = "";
            string strPrevCRMM = "";
            string strIDMS = "";

            if (this.cboYear.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Please select a Year.", "Update Data Required", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            else
            {
                strYear = this.cboYear.SelectedItem.ToString();
            }

            if (this.cboPrevCRMM.SelectedIndex == -1)
            {
                MessageBox.Show("Please select the previous CRMM Date.", "Update Data Required", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            else
            {
                strPrevCRMM = this.cboPrevCRMM.SelectedValue.ToString();
            }

            if ((this.txtRisks.Text == "") || (this.txtAvoidanceScores.Text == "")|| (this.txtCRMMData.Text == ""))
            {
                MessageBox.Show("Please complete the file locations.", "Update Data Required", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (this.txtIDMSData.Text == "")
            {
                var IDMSResult = MessageBox.Show("Do you wish to procced without IDMS Data?", "IDMS Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (IDMSResult == MessageBoxResult.Yes)
                {
                    //proceed with update
                    strIDMS = "No";
                }
                else
                {
                    return;
                }
            }
            else
            {
                strIDMS = "Yes";
            }

            var ProceedResult = MessageBox.Show("Do you wish to proceed with the update?", "Update Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (ProceedResult == MessageBoxResult.Yes)
            {
                //proceed with update
            }
            else
            {
                return;
            }

            // Delete data from temporary files and load data from file locations
            if (importdatafromexcel(this.txtRisks.Text.ToString(), "select * from [sheet1$] where [UTR] is not null", "tmpRisks", "Risks") == false)
            {
                MessageBox.Show("Risk Data Import failed.", "Quarterly Import Process", MessageBoxButton.OK, MessageBoxImage.Exclamation  );
                return;
            }
            if (importdatafromexcel(this.txtAvoidanceScores.Text.ToString(), "select [UTR], [AV score] from [sheet1$] where [UTR] is not null", "tmpAvoidance_Scores", "Avoidance Scores") == false)
            {
                MessageBox.Show("Avoidance Score Data Import failed.", "Quarterly Import Process", MessageBoxButton.OK, MessageBoxImage.Exclamation  );
                return;
            }

            if (importdatafromexcel(this.txtCRMMData.Text.ToString(), "select [UTR], [CRMM_Open_risks], [CRMM_Settled_risks], [CRMM_Highest_settlement] from [sheet1$] where [UTR] is not null", "tmpCRMM", "CRMM") == false)
            {
                MessageBox.Show("CRMM Data Import failed.", "Quarterly Import Process", MessageBoxButton.OK, MessageBoxImage.Exclamation  );
                return;
            }

            if (this.txtIDMSData.Text != "")
            {
                if (importdatafromexcel(this.txtIDMSData.Text.ToString(), "select [UTR], [Open_IDMS], [Closed_IDMS] from [sheet1$] where [UTR] is not null", "tmpIDMS", "IDMS") == false)
                {
                    MessageBox.Show("IDMS Data Import failed.", "Quarterly Import Process", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
            }


            //Run SQL Update Process
            if (ImportPartA(strYear, strPrevCRMM, strIDMS) == false)
            {
                MessageBox.Show("Quarterly Database Update failed.", "Quarterly Import Process", MessageBoxButton.OK, MessageBoxImage.Exclamation  );
                return;
            }

            //confirm import finalised
            MessageBox.Show("Update Completed", "Quarterly Import Process", MessageBoxButton.OK, MessageBoxImage.Information  );
        }
        public bool ImportPartA(string strYear, string strPrevCRMM, string strIDMS)
        {
            string strProcess = "qryQTR_Process_Part";
            StringBuilder errorMessages = new StringBuilder();
            DateTime CRMMdatetime = DateTime.Parse(strPrevCRMM);

            for (int process = 1; process < 4; process++)
            {
                try
                {
                    SqlConnection con = new SqlConnection(Global.ConnectionString);
                    SqlCommand cmd = new SqlCommand(strProcess + process, con);
                    if (process == 2)
                    {
                        cmd.CommandTimeout = 0;
                    }
                    else
                    {
                        cmd.CommandTimeout = Global.TimeOut;
                    }
                    cmd.CommandType = CommandType.StoredProcedure;

                    con.Open();

                    cmd.Parameters.Clear();
                    cmd.CommandText = strProcess + process;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@nYear", strYear);
                    cmd.Parameters.AddWithValue("@nPrevCRMMDate", strPrevCRMM);
                    cmd.Parameters.AddWithValue("@nTodaysDate", DateTime.Now.ToString());
                    cmd.Parameters.AddWithValue("@nIDMS", strIDMS);
                    cmd.ExecuteNonQuery();

                    con.Close();
                    cmd.Dispose();
                    cmd.Parameters.Clear();

                    //return true;
                }
                catch (SqlException ex)
                {
                    for (int i = 0; i < ex.Errors.Count; i++)
                    {
                        errorMessages.Append("Index #" + i + "\n" +
                            "Message: " + ex.Errors[i].Message + "\n" +
                            "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                            "Source: " + ex.Errors[i].Source + "\n" +
                            "Procedure: " + ex.Errors[i].Procedure + "\n");
                    }

                    MessageBox.Show(errorMessages.ToString(), Global.ApplicationName + " - SQL Error : Process_Part" + process, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), Global.ApplicationName + " - Error  Process_Part" + process, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return false;
                }
            }

            return true;
        }
        
        private void PopulateCombos()
        {
            //cboYear
            int intCurrentYear = DateTime.Today.Year;
            for (int intYear = intCurrentYear - 2; intYear <= intCurrentYear; intYear++)
            {
                cboYear.Items.Add(intYear);
            }
            cboYear.SelectedValue = intCurrentYear;

            //cboPrevCRMM
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryGetPrevCRMMDates";
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);

                        DataSet ds = new DataSet();
                        adapter.Fill(ds);

                        con.Close();

                        cboPrevCRMM.ItemsSource = ds.Tables[0].DefaultView;
                        cboPrevCRMM.DisplayMemberPath = "PrevCRMMDates";
                        cboPrevCRMM.SelectedValuePath = "PrevCRMMDates";
                        cboPrevCRMM.SelectedIndex = -1;

                    }
                }
            }
            catch
            {

            }
        }

        private void btnIDMSData_Click(object sender, RoutedEventArgs e)
        {
            this.txtIDMSData.Text = PickFile("Select IDMS file");
        }
    }
}
