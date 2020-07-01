using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;



namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for StandardReports.xaml
    /// </summary>
    public partial class StandardReports : System.Windows.Window
    {
        public StandardReports()
        {
            InitializeComponent();
            PopulateCombos();
            PopulateReports();
        }

        private void PopulateCombos()
        {


            //cboPopulation
            SqlConnection conn = new SqlConnection(Global.ConnectionString);
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM dbo.vwUserPermissionLevel WHERE PID = " + Global.PID + "AND Pop_Friendly_Name IS NOT NULL", conn);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.tblPopulations");
            cboPop.ItemsSource = ds.Tables[0].DefaultView;
            cboPop.DisplayMemberPath = ds.Tables[0].Columns["Pop_Friendly_Name"].ToString();
            cboPop.SelectedValuePath = ds.Tables[0].Columns["Pop_Code_Name"].ToString();
            cboPop.SelectedIndex = 0;

            //cboOffice
            if (Global.AccessLevel == "National")
            {
                SqlDataAdapter sqlda = new SqlDataAdapter("SELECT DISTINCT Office FROM dbo.tblOfficeCRM UNION SELECT 'All'", conn);
                DataSet dset = new DataSet();
                sqlda.Fill(dset, "OfficeList");
                cboOffice.ItemsSource = dset.Tables[0].DefaultView;
                cboOffice.DisplayMemberPath = dset.Tables[0].Columns["Office"].ToString();
                cboOffice.SelectedIndex = 0;
            }
            else
            {
                cboOffice.Items.Add(Global.AccessLevel);
                cboOffice.SelectedIndex = 0;
            }
        }

        private void PopulateReports()
        {
            // Connecting the SQL Server
            SqlConnection con = new SqlConnection(Global.ConnectionString);
            con.Open();
            // Calling the Stored Procedure
            SqlCommand cmd = new SqlCommand("qryGetStndReports", con);
            cmd.CommandType = CommandType.StoredProcedure;


            SqlDataAdapter da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable("dgStReports");
            da.Fill(dt);


            //adding columns to DataTable
            dt.Columns.Add("RunReport", typeof(string)).SetOrdinal(1);
            dt.Columns.Add("ReportYear", typeof(string)).SetOrdinal(2);

            foreach (DataRow dr in dt.Rows)
            {
                if (dr["Year"].ToString() == "True")
                {
                    dr["ReportYear"] = "yyyy";
                }
                else
                {
                    dr["ReportYear"] = "n/a";
                }

                //if (dr["PrevYears"].ToString() == "True")
                //{
                //    dr["Prev_Years"] = "yyyy";
                //}
                //else
                //{
                //    dr["Prev_Years"] = "n/a";
                //}
            }

                dgStReports.ItemsSource = dt.DefaultView;
        }



        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }



        private void cboOffice_DropDownClosed(object sender, EventArgs e)
        {
            string strOffice; /*= ((DataRowView)cboOffice.SelectedItem).Row.ItemArray[0].ToString();*/
            string strPopulation = cboPop.SelectedValue.ToString();

            if (cboOffice.Items.Count == 1)
            {
                strOffice = cboOffice.SelectedItem.ToString();
                
            }
            else
            {
                strOffice = ((DataRowView)cboOffice.SelectedItem).Row.ItemArray[0].ToString();
            }

            SqlConnection conn = new SqlConnection(Global.ConnectionString);
            SqlDataAdapter da = new SqlDataAdapter("SELECT [Team Identifier] FROM dbo.tblOfficeCRM WHERE Office = '" + strOffice + "'AND Pop = '" + strPopulation + "'UNION SELECT 'All'", conn);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.tblOfficeCRM");
            cboTeam.ItemsSource = ds.Tables[0].DefaultView;
            cboTeam.DisplayMemberPath = ds.Tables[0].Columns["Team Identifier"].ToString();
            cboTeam.SelectedIndex = 0;


        }

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.CheckBox check = sender as System.Windows.Controls.CheckBox;
            DataRowView selectedRow = (DataRowView)dgStReports.SelectedItem;


            //open dialog box to select the year
            if ((selectedRow["ReportYear"].ToString() == "yyyy") && (check.IsChecked == true))
            {
                Year year = new Year();
                year.ShowDialog();

                selectedRow["ReportYear"] = Year.strReportYear;
            }

            //update row if the report has to be generated
            if (check.IsChecked == true)
                {selectedRow["RunReport"] = "Yes";}
            else
                {selectedRow["RunReport"] = "";}
        }

        private void btnReports_Click(object sender, RoutedEventArgs e)
        {

            string strPopulation = cboPop.SelectedValue.ToString();
            string strOffice = cboOffice.Text.ToString();
            string strTeam = cboTeam.Text.ToString();
            int intYear;

            foreach (DataRowView row in dgStReports.Items)
            {
                if (row["RunReport"].ToString() == "Yes")
                {
                    //MessageBox.Show(
                    //    cboPop.Text.ToString() +"\n"+
                    //    cboOffice.Text.ToString() +"\n"+
                    //    cboTeam.Text.ToString() +"\n"+
                    //    row["ReportYear"].ToString() +"\n"+
                    //    row["Template"].ToString() +"\n"+
                    //    row["ExcelReportName"].ToString());

                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                    System.Data.DataTable dt = new System.Data.DataTable("Report");

                    string StoredProcedureName = "qry" + row["ExcelReportName"].ToString(); //Main SQL Stored Procedure

                    try
                    {
                        intYear = Convert.ToInt32(row.Row["ReportYear"]);
                    }
                    catch
                    {
                        intYear = 0;//Year is n/a
                    }
                    

                    // Connecting the SQL Server
                    SqlConnection con = new SqlConnection(Global.ConnectionString);
                    con.Open();
                    // Calling the Stored Procedure
                    SqlCommand cmd = new SqlCommand(StoredProcedureName, con);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@nPop", SqlDbType.Text).Value = strPopulation;
                    cmd.Parameters.Add("@nOffice", SqlDbType.Text).Value = strOffice;
                    cmd.Parameters.Add("@nTeam", SqlDbType.Text).Value = strTeam;
                    cmd.Parameters.Add("@nYear", SqlDbType.Int).Value = intYear;


                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);

                    dt = ds.Tables[0];

                    /*Set up work book, work sheets, and excel application*/
                    Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook obook = null;
                    Microsoft.Office.Interop.Excel._Worksheet osheet = null;
                    //try
                    //{
                    //    string path = AppDomain.CurrentDomain.BaseDirectory;
                    object misValue = System.Reflection.Missing.Value;
                    //    if (strFileName == "")
                    //    {
                    // Template hasn't been specified so use a new blank workbook
                    obook = oexcel.Workbooks.Add(misValue);
                    osheet = (Microsoft.Office.Interop.Excel.Worksheet)obook.Sheets["Sheet1"];
                    //    }
                    //    else
                    //    {
                    //        // Template was specified so open the specified template
                    //        strFileName = path + strFileName;
                    //        obook = oexcel.Workbooks.Open(strFileName);
                    //        osheet = (Microsoft.Office.Interop.Excel.Worksheet)obook.Sheets["Data"];
                    //        osheet.UsedRange.ClearContents();
                    //    }

                    int colIndex = 0;
                    int rowIndex = 1;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        osheet.Cells[1, colIndex] = dc.ColumnName;
                    }
                    foreach (DataRow dr in dt.Rows)
                    {
                        rowIndex++;
                        colIndex = 0;

                        foreach (DataColumn dc in dt.Columns)
                        {
                            colIndex++;
                            osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                        }
                    }

                    osheet.Columns.AutoFit();
                    //    }
                    //catch (Exception ex)
                    //{

                    //}

                    //this.TopMost = false;

                    oexcel.Visible = true;

                    //releaseObject(osheet);

                    //releaseObject(obook);

                    //releaseObject(oexcel);

                    GC.Collect();

                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                }
            }
            MessageBox.Show("Report function has finished running.","Wealthy RPT",MessageBoxButton.OK,MessageBoxImage.Information);
            this.Close();
        }

    }
}
