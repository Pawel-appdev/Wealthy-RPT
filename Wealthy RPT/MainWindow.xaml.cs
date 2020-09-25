using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using Ini;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public int intPageSize = LoadAppVariables.MainPageSize;// number of rows shown on datagrid page
        public int intOffset;
        public int intCurrentPage = 1;
        public string strMove;
        public int intTotalRows;
        public int intLastPage;
        public bool blnCombosInitialised = false;

        #region menu items
        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.A && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                mnuSimpleReports_Click(sender, e);
            }
            if (e.Key == Key.S && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                mnuStandardReports_Click(sender, e);
            }
        }
        private void MnuExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
            Environment.Exit(0);
        }

        private void MnuAbout_Click(object sender, RoutedEventArgs e)
        {
            About about = new About();
            about.Show();

            //SplashScreen about = new SplashScreen();
            //about.Show();
        }

        private void MnuKey_Click(object sender, RoutedEventArgs e)
        {
            Key_Legend legend = new Key_Legend();
            legend.Show();
        }

        private void MnuNewCase_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                RPT_Detail rPT_Detail = new RPT_Detail();
                User user = new User();
                Globals.blnIgnoreEvents = true;

                rPT_Detail.cboPopCode.SelectedValue = user.Pop_Code_Name.ToString();
                rPT_Detail.cboPopFriendly.SelectedIndex = rPT_Detail.cboPopCode.SelectedIndex;

                rPT_Detail.cboSegment.IsReadOnly = false;
                rPT_Detail.cboSegment.IsEnabled = true;
                rPT_Detail.cboSegment.IsTabStop = true;

                Globals.blnIgnoreEvents = false;
                rPT_Detail.lblPopYear.Text = DateTime.Now.Year.ToString();
                rPT_Detail.tcmdSave.Text = "Add";
                rPT_Detail.tcmdSave.HorizontalAlignment = HorizontalAlignment.Center;
                rPT_Detail.tcmdSave.ToolTip = "Add";
                rPT_Detail.cmdUpdateClose.Visibility = Visibility.Hidden;

                rPT_Detail.Show();
            }
            catch
            {

            }
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

        private void mnuStandardReports_Click(object sender, RoutedEventArgs e)
        {
            StandardReports standardReports = new StandardReports();
            standardReports.Show();
        }

        private void mnuSimpleReports_Click(object sender, RoutedEventArgs e)
        {
            SimpleReports simpleReports = new SimpleReports();
            simpleReports.Show();
        }

        private void mnuImportData_Click(object sender, RoutedEventArgs e)
        {
            ImportData importData = new ImportData();
            importData.Show();
        }
        #endregion
        
        public MainWindow()
        {
            InitializeComponent();
            PopulateCombos();
            blnCombosInitialised = true;

            RAG.GetRAGBreaks();

            txtRows.Text = intPageSize.ToString();

            this.sbiPID.Content = "PID : " + Global.PID;
            this.sbiName.Content = "Name : " + Global.FullName;
            this.sbiVersion.Content = "Version : " + Global.CurrentVersion;
            this.sbiAccessLevel.Content = "Access Level : " + Global.AccessLevel;

            HideMenuItems();
            PopulateCases();
        }

        private void HideMenuItems()
        {
            if (Global.Admin != true)
            {
                mnuAdmin.IsEnabled = false;
            }
        }
        private void PopulateCases()
        {
            int intYear = Convert.ToInt32(cboYear.SelectedValue);
            string strOffice = (cboOffice.SelectedIndex == -1) ? "All" : cboOffice.SelectedValue.ToString();
            string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : cboTeam.SelectedValue.ToString();
            int intPID = Convert.ToInt32(Global.PID);
            string strPop = "";
            try
            {
                strPop = this.cboPopulation.SelectedValue.ToString();
            }
            catch
            {

            }

            intOffset = 0;

            GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
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

            //cboPopulation
            SqlConnection conn = new SqlConnection(Global.ConnectionString);
            //SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM dbo.tblPopulations", conn);//get all populations
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM dbo.vwUserPermissionLevel WHERE PID = " + Global.PID + "AND Pop_Friendly_Name IS NOT NULL", conn);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.tblPopulations");
            cboPopulation.ItemsSource = ds.Tables[0].DefaultView;
            cboPopulation.DisplayMemberPath = ds.Tables[0].Columns["Pop_Friendly_Name"].ToString();
            cboPopulation.SelectedValuePath = ds.Tables[0].Columns["Pop_Code_Name"].ToString();
            cboPopulation.SelectedIndex = 0;

            //cboOffice
            if(Global.AccessLevel == "National")
            {
                SqlDataAdapter sqlda = new SqlDataAdapter("SELECT DISTINCT Office FROM dbo.tblOfficeCRM UNION SELECT 'All'", conn);
                DataSet dset = new DataSet();
                sqlda.Fill(dset, "OfficeList");
                cboOffice.ItemsSource = dset.Tables[0].DefaultView;
                cboOffice.DisplayMemberPath = dset.Tables[0].Columns["Office"].ToString();
            }
            else
            {
                cboOffice.Items.Add(Global.AccessLevel);
            }
        }

        private void rbMyCases_Checked(object sender, RoutedEventArgs e)
        {
            groupOfficeTeam.Visibility = Visibility.Hidden ;

            int intPID = Convert.ToInt32(Global.PID);
            int intYear;
            string strPop;
            string strOffice = (cboOffice.SelectedIndex == -1) ? "All" : cboOffice.SelectedValue.ToString();
            string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : cboTeam.SelectedValue.ToString();
            try
            {
                if (string.IsNullOrEmpty(this.cboYear.Text))
                {
                    intYear = DateTime.Today.Year;
                }
                else
                {
                    intYear = Convert.ToInt32(this.cboYear.SelectedValue);
                }
                strPop = cboPopulation.SelectedValue.ToString();
            }
            catch
            {
                intYear = DateTime.Today.Year;
                strPop = "rPt10Mill";
            }
            
            resetValues();

            if (blnCombosInitialised == true)
            {
                GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
            }
        }

        private void rbOtherCases_Checked(object sender, RoutedEventArgs e)
        {
            string strPop = "";
            try
            {
            strPop = this.cboPopulation.SelectedValue.ToString();
            }
            catch
            {
                MessageBox.Show("Unable to establish your access level. The application will now close.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                Environment.Exit(0);
            }

            groupOfficeTeam.Visibility = Visibility.Visible;
            clearOtherCasesCombos();
            int intYear = Convert.ToInt32(cboYear.SelectedValue);
            string strOffice = (cboOffice.SelectedIndex == -1) ? "All" : cboOffice.SelectedValue.ToString();
            string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : cboTeam.SelectedValue.ToString();
            int intPID = Convert.ToInt32(Global.PID);

            resetValues();

            GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
        }

        private void cboYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int intYear = Convert.ToInt32(cboYear.SelectedValue);
                string strOffice = (cboOffice.SelectedIndex == -1) ? "All" : ((DataRowView)cboOffice.SelectedItem).Row.ItemArray[0].ToString();
                string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : ((DataRowView)cboTeam.SelectedItem).Row.ItemArray[0].ToString();
                int intPID = Convert.ToInt32(Global.PID);
                string strPop = this.cboPopulation.SelectedValue.ToString();

                resetValues();

                GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
            }
            catch
            {
                //the value of cboPopulation has not been selected yet - do nothing
            }
        }
        private void cboPopulation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intYear = Convert.ToInt32(cboYear.SelectedValue);
            string strOffice;

            if (cboOffice.Items.Count == 1)//user allowed to see cases fron 1 office
            {
                strOffice = Global.AccessLevel.ToString();
            }
            else
            {
                strOffice = (cboOffice.SelectedIndex == -1) ? "All" : ((DataRowView)cboOffice.SelectedItem).Row.ItemArray[0].ToString();
            }

            string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : ((DataRowView)cboTeam.SelectedItem).Row.ItemArray[0].ToString();
            int intPID = Convert.ToInt32(Global.PID);
            string strPop = cboPopulation.SelectedValue.ToString();

            resetValues();

            if (blnCombosInitialised == true)
            {
                GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
            }
        }


        private void cboOffice_DropDownClosed(object sender, EventArgs e)
        {
            try//to get teams for specific office
            {  

                int intYear = Convert.ToInt32(cboYear.SelectedValue);
                string strOffice;
                if(cboOffice.Items.Count == 1)
                {
                    strOffice = cboOffice.SelectedItem.ToString();
                }
                else
                {
                    strOffice = ((DataRowView)cboOffice.SelectedItem).Row.ItemArray[0].ToString();
                }
                
                string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : cboTeam.SelectedValue.ToString();

                if(strTeam== "System.Data.DataRowView"){strTeam = "All";}//user previously selected Team(SelectedIndex <> -1, wants to see Office Cases
                int intPID = Convert.ToInt32(Global.PID);
                string strPopulation = cboPopulation.SelectedValue.ToString();
                SqlConnection conn = new SqlConnection(Global.ConnectionString);
                SqlDataAdapter da = new SqlDataAdapter("SELECT [Team Identifier] FROM dbo.tblOfficeCRM WHERE Office = '" + strOffice + "'AND Pop = '" + strPopulation + "'UNION SELECT 'All'", conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "dbo.tblOfficeCRM");
                cboTeam.ItemsSource = ds.Tables[0].DefaultView;
                cboTeam.DisplayMemberPath = ds.Tables[0].Columns["Team Identifier"].ToString();

                string strPop = cboPopulation.SelectedValue.ToString();
                resetValues();

                GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
        }
            catch
            {
            //    MessageBox.Show("Unable to get cases from this office.", "Error",MessageBoxButton.OK, MessageBoxImage.Error);
            //    cboOffice.Text = "";
            }
        }

        private void cboTeam_DropDownClosed(object sender, EventArgs e)
        {

            try
            {
                int intYear = Convert.ToInt32(cboYear.SelectedValue);
                string strOffice = cboOffice.Text.ToString();
                string strTeam = ((DataRowView)cboTeam.SelectedItem).Row.ItemArray[0].ToString();
                int intPID = Convert.ToInt32(Global.PID);
                string strPop = cboPopulation.SelectedValue.ToString();

                resetValues();

                GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
            }
            catch
            {
                //MessageBox.Show("Unable to get cases from this team.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //cboTeam.Text = "";
            }
        }


        private void GetdgCases(int intYear, string strOffice, string strTeam, int intPID, string strPop, int intOffset)
        {
            try
            {

                if (Global.AccessLevel != "National")//user allowed to see cases from 1 office
            { 
                strOffice = Global.AccessLevel;
                cboOffice.Text = strOffice;
            }

                DataTable dt = new DataTable("dgCases");


            if (rbMyCases.IsChecked == true)//========== Get My Cases ==========
            {
                // Connecting the SQL Server
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                con.Open();
                // Calling the Stored Procedure
                SqlCommand cmd = new SqlCommand("qryGetMyCases", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@nYear", SqlDbType.Int).Value = intYear;
                cmd.Parameters.Add("@nPID", SqlDbType.Int).Value = intPID;
                cmd.Parameters.Add("@nPop", SqlDbType.Text).Value = strPop;
                cmd.Parameters.Add("@nRows", SqlDbType.Int).Value = intPageSize ;
                cmd.Parameters.Add("@nOffset", SqlDbType.Int).Value = intOffset;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);

                try//show My Cases
                {
                    this.txtPage.Text = "Showing page 0 of 0";
                    this.dgCases.ItemsSource = dt.DefaultView;
                    ////hiding Office , Team & Year columns
                    this.dgCases.Columns[3].Visibility = Visibility.Hidden;
                    this.dgCases.Columns[4].Visibility = Visibility.Hidden;
                    this.dgCases.Columns[5].Visibility = Visibility.Hidden;

                        foreach (DataRow dr in dt.Rows)
                        {
                            String strPopulation = dr["Pop"].ToString().Replace(" ", string.Empty);

                            if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == true)
                            {
                                dr["DailyRank"] = 0;
                            }
                            else
                            {
                                dr["DailyRank"] = decimal.Round(Convert.ToDecimal(dr["DailyRank"]), 2, MidpointRounding.AwayFromZero);
                            }
                        }

                    // second table
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    intTotalRows = Convert.ToInt32(ds.Tables[1].Rows[0]["NumberOfRows"]);

                    updatePageCounter();
                }
                catch { /*MessageBox.Show("Unable to retrieve all data")*//*txtPage.Text = "Showing page 0 of 0"*/; }//no My Cases found
                }
            else//                          ========== Get Other Cases ==========
            {
                if (strOffice == "All" && strTeam == "All")//   Get National Cases
                {
                    // Connecting the SQL Server
                    SqlConnection con = new SqlConnection(Global.ConnectionString);
                    con.Open();
                    // Calling the Stored Procedure
                    SqlCommand cmd = new SqlCommand("qryGetNationalCases", con);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@nYear", SqlDbType.Int).Value = intYear;
                    cmd.Parameters.Add("@nPID", SqlDbType.Int).Value = intPID;
                    cmd.Parameters.Add("@nPop", SqlDbType.Text).Value = strPop;
                    cmd.Parameters.Add("@nRows", SqlDbType.Int).Value = intPageSize;
                    cmd.Parameters.Add("@nOffset", SqlDbType.Int).Value = intOffset;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(dt);

                    // second table
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    intTotalRows = Convert.ToInt32(ds.Tables[1].Rows[0]["NumberOfRows"]);

                    updatePageCounter();

                }
                if (strOffice != "All" && strTeam == "All")//   Get Office Cases
                {
                    // Connecting the SQL Server
                    SqlConnection con = new SqlConnection(Global.ConnectionString);
                    con.Open();
                    // Calling the Stored Procedure
                    SqlCommand cmd = new SqlCommand("qryGetOfficeCases", con);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@nYear", SqlDbType.Int).Value = intYear;
                    cmd.Parameters.Add("@nOffice", SqlDbType.Text).Value = strOffice;
                    cmd.Parameters.Add("@nPID", SqlDbType.Int).Value = intPID;
                    cmd.Parameters.Add("@nPop", SqlDbType.Text).Value = strPop;
                    cmd.Parameters.Add("@nRows", SqlDbType.Int).Value = intPageSize;
                    cmd.Parameters.Add("@nOffset", SqlDbType.Int).Value = intOffset;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(dt);

                    // second table
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    intTotalRows = Convert.ToInt32(ds.Tables[1].Rows[0]["NumberOfRows"]);

                    updatePageCounter();

                }
                if (strOffice != "All" && strTeam != "All")//   Get Team Cases
                {
                    // Connecting the SQL Server
                    SqlConnection con = new SqlConnection(Global.ConnectionString);
                    con.Open();
                    // Calling the Stored Procedure
                    SqlCommand cmd = new SqlCommand("qryGetTeamCases", con);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@nYear", SqlDbType.Int).Value = intYear;
                    cmd.Parameters.Add("@nOffice", SqlDbType.Text).Value = strOffice;
                    cmd.Parameters.Add("@nTeam", SqlDbType.Text).Value = strTeam;
                    cmd.Parameters.Add("@nPID", SqlDbType.Int).Value = intPID;
                    cmd.Parameters.Add("@nPop", SqlDbType.Text).Value = strPop;
                    cmd.Parameters.Add("@nRows", SqlDbType.Int).Value = intPageSize;
                    cmd.Parameters.Add("@nOffset", SqlDbType.Int).Value = intOffset;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(dt);

                    // second table
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    intTotalRows = Convert.ToInt32(ds.Tables[1].Rows[0]["NumberOfRows"]);

                    updatePageCounter();
                }




            }

                // ##############
                //unhiding Office & Team columns
                dgCases.Columns[3].Visibility = Visibility.Hidden;
                dgCases.Columns[4].Visibility = Visibility.Visible;
                dgCases.Columns[5].Visibility = Visibility.Visible;

                //adding columns to DataTable
                dt.Columns.Add("SegNo", typeof(int)).SetOrdinal(1);
                dt.Columns.Add("ProSegNo", typeof(int)).SetOrdinal(2);
                dt.Columns.Add("SegMove", typeof(int)).SetOrdinal(3);


                foreach (DataRow dr in dt.Rows)
                {
                    String strPopulation = dr["Pop"].ToString().Replace(" ", string.Empty);

                    if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == true)
                    {
                        dr["DailyRank"] = 0;
                    }
                    else
                    {
                        dr["DailyRank"] = decimal.Round(Convert.ToDecimal(dr["DailyRank"]), 2, MidpointRounding.AwayFromZero);
                    }

                    if (string.IsNullOrEmpty(dr["Strand"].ToString()) == true)
                    {
                        dr["Strand"] = "";
                    }
                    else
                    {
                        switch (dr["Strand"].ToString())
                        {

                            case "High Risk Wealth":
                                dr["Strand"] = "HRW";
                                break;
                            case "Tax Events & Assurance":
                                dr["Strand"] = "TEA";
                                break;
                        }
                    }

                    if (Global.DisplayRAG == "True")
                    {
                        switch (strPopulation)
                        {

                            case "rPt10Mill":
                                if (string.IsNullOrEmpty(dr["Segment"].ToString()) == true)
                                {
                                    dr["SegNo"] = 0;
                                    dr["Segment"] = "NYR";
                                }


                                if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == false)
                                {

                                    if (Convert.ToDouble(dr["DailyRank"]) <= RAG.RAG10M_1)
                                    {
                                        dr["Segment"] = "Cert";
                                        dr["SegNo"] = 1;
                                    }
                                    else if (Convert.ToDouble(dr["DailyRank"]) > RAG.RAG10M_2)
                                    {
                                        dr["Segment"] = "High";
                                        dr["SegNo"] = 3;
                                    }
                                    else
                                    {
                                        dr["Segment"] = "Res";
                                        dr["SegNo"] = 2;
                                    }
                                }
                                break;

                            case "rPt20Mill":
                                if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == false)
                                {

                                    if (Convert.ToDouble(dr["DailyRank"]) <= RAG.RAG20M_1)
                                    {
                                        dr["Segment"] = "Cert";
                                        dr["SegNo"] = 1;
                                    }
                                    else if (Convert.ToDouble(dr["DailyRank"]) > RAG.RAG20M_2)
                                    {
                                        dr["Segment"] = "High";
                                        dr["SegNo"] = 3;
                                    }
                                    else
                                    {
                                        dr["Segment"] = "Res";
                                        dr["SegNo"] = 2;
                                    }
                                }
                                break;

                            default:
                                if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == false)
                                {

                                    if (Convert.ToDouble(dr["DailyRank"]) <= 33.33)
                                    {
                                        dr["Segment"] = "Cert";
                                        dr["SegNo"] = 1;
                                    }
                                    else if (Convert.ToDouble(dr["DailyRank"]) > 66.67)
                                    {
                                        dr["Segment"] = "High";
                                        dr["SegNo"] = 3;
                                    }
                                    else
                                    {
                                        dr["Segment"] = "Res";
                                        dr["SegNo"] = 2;
                                    }
                                }
                                break;

                        }
                    }
                    else//do not show RAG Status
                    {

                        switch (strPopulation)
                        {

                            case "rPt10Mill":
                                if (string.IsNullOrEmpty(dr["Segment"].ToString()) == true)
                                {
                                    dr["SegNo"] = 0;
                                    dr["Segment"] = "NYR";
                                }


                                if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == false)
                                {

                                    if (Convert.ToDouble(dr["DailyRank"]) <= RAG.RAG10M_1)
                                    {
                                        //dr["Segment"] = "Cert";
                                        dr["SegNo"] = 1;
                                    }
                                    else if (Convert.ToDouble(dr["DailyRank"]) > RAG.RAG10M_2)
                                    {
                                        //dr["Segment"] = "High";
                                        dr["SegNo"] = 3;
                                    }
                                    else
                                    {
                                        //dr["Segment"] = "Res";
                                        dr["SegNo"] = 2;
                                    }
                                }
                                break;

                            case "rPt20Mill":
                                if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == false)
                                {

                                    if (Convert.ToDouble(dr["DailyRank"]) <= RAG.RAG20M_1)
                                    {
                                        //dr["Segment"] = "Cert";
                                        dr["SegNo"] = 1;
                                    }
                                    else if (Convert.ToDouble(dr["DailyRank"]) > RAG.RAG20M_2)
                                    {
                                        //dr["Segment"] = "High";
                                        dr["SegNo"] = 3;
                                    }
                                    else
                                    {
                                        //dr["Segment"] = "Res";
                                        dr["SegNo"] = 2;
                                    }
                                }
                                break;

                            default:
                                if (string.IsNullOrEmpty(dr["DailyRank"].ToString()) == false)
                                {

                                    if (Convert.ToDouble(dr["DailyRank"]) <= 33.33)
                                    {
                                        //dr["Segment"] = "Cert";
                                        dr["SegNo"] = 1;
                                    }
                                    else if (Convert.ToDouble(dr["DailyRank"]) > 66.67)
                                    {
                                        //dr["Segment"] = "High";
                                        dr["SegNo"] = 3;
                                    }
                                    else
                                    {
                                        //dr["Segment"] = "Res";
                                        dr["SegNo"] = 2;
                                    }
                                }
                                break;

                        }

                    }
                    String strProceedingSegment = dr["ProceedingSegment"].ToString().ToUpper();

                    switch (strProceedingSegment)
                    {
                        case "NYR":
                            dr["ProSegNo"] = 0;
                            break;

                        case "CERT":
                            dr["ProSegNo"] = 1;
                            break;

                        case ""://do nothing case has remained in same segment following bulk update
                            dr["ProSegNo"] = 1;
                            break;

                        case "RES":
                            dr["ProSegNo"] = 2;
                            break;

                        case "HIGH":
                            dr["ProSegNo"] = 3;
                            break;
                    }

                    //dgCases.ItemsSource = dt.DefaultView;

                    if (Convert.ToInt16(dr["SegNo"]) > Convert.ToInt16(dr["ProSegNo"])) { dr["SegMove"] = 1; }// fix previous years
                    else if (Convert.ToInt16(dr["SegNo"]) < Convert.ToInt16(dr["ProSegNo"])) { dr["SegMove"] = -1; }
                    else { dr["SegMove"] = 0; }
                }


                // #########
                try//show Other Cases
            {
                this.dgCases.ItemsSource = dt.DefaultView;
            }
            catch { }//no Other Cases found

            }
            catch
            {
                MessageBox.Show("Unable to connect to database.");
            }
        }

        private void clearOtherCasesCombos()
        {
            cboOffice.Text = "";
            cboTeam.Text = "";
            intCurrentPage = 1;
        }

        private void btnFirstPage_Click(object sender, RoutedEventArgs e)
        {
            strMove = "first";
            NavigatePages(strMove);
        }
        private void btnPreviousPage_Click(object sender, RoutedEventArgs e)
        {
            strMove = "prev";
            NavigatePages(strMove);
        }


        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            strMove = "next";
            NavigatePages(strMove);
        }
        private void btnLastPage_Click(object sender, RoutedEventArgs e)
        {
            strMove = "last";
            NavigatePages(strMove);
        }
        private void NavigatePages(string strMove)
        {
            int intFirstPage = 1;

            switch (strMove)
            {
                case "first":
                    resetValues();
                    break;

                case "prev":
                    if (intCurrentPage == intFirstPage || intOffset == 0) { return; }
                    intOffset = intOffset - intPageSize;
                    intCurrentPage = --intCurrentPage;
                    break;

                case "next":
                    if (intCurrentPage == intLastPage) { return; }
                    intOffset = intOffset + intPageSize;
                    intCurrentPage = ++intCurrentPage;
                    break;

                case "last":
                    //decrease intOffset by intPageSize othervise last page would be empty
                    if( (intTotalRows % intPageSize == 0))
                    {
                        intTotalRows = intTotalRows - intPageSize;
                    }
                    else
                    {
                        //find highest row number where MOD(intPagesize) = 0  
                        while (intTotalRows % intPageSize != 0)
                        {
                            intTotalRows--;
                        }
                    }

                    intOffset = intTotalRows;
                    intCurrentPage = intLastPage;
                    break;
            }

            int intYear = Convert.ToInt32(cboYear.SelectedValue);
            string strOffice = (cboOffice.Text == "") ? "All" : cboOffice.Text.ToString();

            string strTeam = (cboTeam.Text == "") ? "All" : cboTeam.Text.ToString();
            int intPID = Convert.ToInt32(Global.PID);
            string strPop = cboPopulation.SelectedValue.ToString();


            GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
        }

        private void resetValues()
        {
            intOffset = 0;
            intCurrentPage = 1;

            try 
            {
                this.txtSearch.Text = "";
                this.dgCases.Margin = new Thickness(32, 150, 25, 38);

            } 
            catch { }
            
        }
        private void updatePageCounter()
        {
            //updatetxtPage(intTotalRows, intOffset);
            double dblTotalPages;
            dblTotalPages = (double)intTotalRows / (double)intPageSize;
            intLastPage = Convert.ToInt32(Math.Ceiling(dblTotalPages));

            if (dblTotalPages == 0)//nothing to show
            {
                intCurrentPage = 0;
            }

            txtPage.Text = "Showing page " + intCurrentPage.ToString() + " of " + intLastPage.ToString();
        }

        private void btnGoTo_Click(object sender, RoutedEventArgs e)
        {

            if (txtGoTo.Text != "" && txtGoTo.Text.All(char.IsDigit) && Convert.ToInt32(txtGoTo.Text) > 0)
            {
                int intGoToPage = Convert.ToInt32(txtGoTo.Text);
                txtGoTo.Text = "";

                if (intGoToPage > intLastPage) 
                {
                    MessageBox.Show("You can view only " + intLastPage + " page(s).", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    txtGoTo.Text = "";
                    txtGoTo.Focus();
                    return;
                }

                intCurrentPage = intGoToPage;
                intOffset = intCurrentPage * intPageSize - intPageSize;

                int intYear = Convert.ToInt32(cboYear.SelectedValue);
                string strOffice = (cboOffice.SelectedIndex == -1) ? "All" : cboOffice.SelectedValue.ToString();
                string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : cboTeam.SelectedValue.ToString();
                int intPID = Convert.ToInt32(Global.PID);
                string strPop = cboPopulation.SelectedValue.ToString();

                GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
            }
            else
            {
                MessageBox.Show("Please enter a whole numbers only.");
                txtGoTo.Focus();
            }
        }

        private void btnRows_Click(object sender, RoutedEventArgs e)
        {
            if (txtSearch.Text == "")
            {
                resetValues();

                if (txtRows.Text != "" && txtRows.Text.All(char.IsDigit) && Convert.ToInt32(txtRows.Text) > 0)
                {
                    intPageSize = Convert.ToInt32(txtRows.Text);
                    int intMaxPageSize = 1000;
                    if(intPageSize > intMaxPageSize)
                    {
                        MessageBox.Show("Limited to a maximum number of " + intMaxPageSize + " rows.","Rows per page");

                        txtRows.Text = intMaxPageSize.ToString();
                        intPageSize = intMaxPageSize;
                    }
                    int intYear = Convert.ToInt32(cboYear.SelectedValue);
                    string strOffice = (cboOffice.SelectedIndex == -1) ? "All" : cboOffice.SelectedValue.ToString();
                    string strTeam = (cboTeam.SelectedIndex == -1) ? "All" : cboTeam.SelectedValue.ToString();
                    int intPID = Convert.ToInt32(Global.PID);
                    string strPop = cboPopulation.SelectedValue.ToString();

                    intOffset = 0;

                    SaveMainPageSize(intPageSize);
                    GetdgCases(intYear, strOffice, strTeam, intPID, strPop, intOffset);
                }

                else
                {
                    MessageBox.Show("Please enter a whole numbers only.");
                    txtRows.Focus();
                }
            }
            else
            {
                intPageSize = Convert.ToInt32(txtRows.Text);
                KeywordSearch();
            }

        }

        private void SaveMainPageSize(int intPageSize)
        {
            IniFile LocalFile = new IniFile(LoadAppVariables.LocalFile);
            LocalFile.IniWriteValue("UserSettings", "MainPageSize", Convert.ToString(intPageSize));
        }

            private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(txtSearch.Text.Length > 2)
            {
                KeywordSearch();
            }
            else
            {
                MessageBox.Show("Search box must contain at least three characters.", "Try again", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            
        }
        private void KeywordSearch()
        {
            int intPopulations = cboPopulation.Items.Count;

            if (intPopulations == 0)
            {
                MessageBox.Show("Unable to establish your access level. The application will now close.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                Environment.Exit(0);
            }

            dgCases.Columns[4].Visibility = Visibility.Visible;
            dgCases.Columns[5].Visibility = Visibility.Visible;

            DataTable dt = new DataTable("Search");

            // Connecting the SQL Server
            SqlConnection con = new SqlConnection(Global.ConnectionString);
            con.Open();
            // Calling the Stored Procedure
            SqlCommand cmd = new SqlCommand("qrySearch_UTR_Name", con);
            cmd.CommandType = CommandType.StoredProcedure;


            cmd.Parameters.Add("@SearchStr", SqlDbType.Text).Value = txtSearch.Text;
            cmd.Parameters.Add("@AccessLevel", SqlDbType.Text).Value = Global.AccessLevel.ToString();
            cmd.Parameters.Add("@nPID", SqlDbType.Int).Value = Global.PID;
            //cmd.Parameters.Add("@nRows", SqlDbType.Int).Value = intPageSize;
            //cmd.Parameters.Add("@nOffset", SqlDbType.Int).Value = intOffset;

            SqlDataAdapter da = new SqlDataAdapter(cmd);

            da.Fill(dt);

            //// second table
            //DataSet ds = new DataSet();
            //da.Fill(ds);
            //intTotalRows = Convert.ToInt32(ds.Tables[1].Rows[0]["NumberOfRows"]);

            intTotalRows = 1;
            intLastPage = 1;
            txtRows.Text = "";
            updatePageCounter();

            dgCases.ItemsSource = dt.DefaultView;
            dgCases.Columns[3].Visibility = Visibility.Visible;

            rbMyCases.IsChecked = false;
            rbOtherCases.IsChecked = false;
            dgCases.Margin = new Thickness(32, 100, 25, 38);
        }

        private void mnuPercCalculation_Click(object sender, RoutedEventArgs e)
        {
            
           IniFile GlobalFile = new IniFile(LoadAppVariables.GlobalFile);


            GlobalFile.IniWriteValue("System", "DailyRecalc", DateTime.Now.ToString("dd/MM/yyyy"));
            
        }

        private void DgCases_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            RPT_Detail rptDetail = new RPT_Detail();  // initialise form
            RPT.RPT_Data rpt = new RPT.RPT_Data(); // initialise data
            // get data for selected UTR
            double dUTR = 0;
            try { dUTR = Convert.ToDouble((dgCases.Columns[1].GetCellContent(dgCases.CurrentCell.Item) as TextBlock).Text); } catch { }
            if (dUTR == 0) // handles user double-clicks outside populated rows
            {
                return;
            }
            int iYear = 2000;
            try { iYear = Convert.ToInt16(cboYear.SelectedValue); } catch { }
            double dPercentile = 0;
            int iPercentileIndex = dgCases.Columns.IndexOf(dgCases.Columns.FirstOrDefault(c => c.Header.ToString() == "Daily Ranking"));
            try { dPercentile = Convert.ToDouble((dgCases.Columns[iPercentileIndex].GetCellContent(dgCases.CurrentCell.Item) as TextBlock).Text); } catch { dPercentile = 0; } // text can be 'N/A'
            string sPop = "";
            try { sPop = this.cboPopulation.SelectedValue.ToString(); } catch {}
            if (rpt.GetRPDData(dUTR, iYear, dPercentile, sPop) == false)
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                rptDetail.Close();
                MessageBox.Show("Problem loading WRT Data.", "WRT Data", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                rptDetail.DataContext = rpt;
                //rptDetail.mscHistory.ItemsSource = Globals.dtGraph.DefaultView;
                //rptDetail.dgHistorical.ItemsSource = Globals.dtGrid.DefaultView;
                rptDetail.cboPopFriendly.SelectedIndex = this.cboPopulation.SelectedIndex;
                rptDetail.cboPopCode.SelectedIndex = this.cboPopulation.SelectedIndex;
                try { rptDetail.lblPopYear.Text = this.cboYear.SelectedValue.ToString(); } catch { DateTime.Now.Year.ToString(); }
                rptDetail.Closing += new System.ComponentModel.CancelEventHandler(Refresh_Cases); /*grab child form closing event for parent refesh */
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                rptDetail.Show();
            }
        }

        public void Refresh_Cases(object sender, System.ComponentModel.CancelEventArgs e)
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            PopulateCases();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

        private void mnuQuarterlyUpdate_Click(object sender, RoutedEventArgs e)
        {
            PeriodImport periodImport = new PeriodImport();
            periodImport.Show();
        }

        private void mnuUpdateDBTables_Click(object sender, RoutedEventArgs e)
        {
            Admin ctr_Admin = new Admin();
            ctr_Admin.ShowDialog();
        }

        private void MnuGuidance_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = Cursors.Wait;
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@Global.GuidanceFile);
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                //xlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                xlApp.Visible = true;
            }
            catch(Exception ex)
            {
                GC.Collect();
                Cursor = Cursors.Arrow;
                MessageBox.Show("Cannot open the Guidance." + Environment.NewLine
                    + Environment.NewLine + ex.Message, "Wealthy Risk Tool", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            Cursor = Cursors.Arrow;
        }
    }
}
