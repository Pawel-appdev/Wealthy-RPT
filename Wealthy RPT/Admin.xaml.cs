using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for Admin.xaml
    /// </summary>
    public partial class Admin : Window
    {
        public static string strAddlDataInstr;
        public static string strMainKey;


        public Admin()
        {
            InitializeComponent();
            PopulateTables();
        }

        private void PopulateTables()
        {
            string strSQLcommand;

            if(Global.BA_Admin == true)
            {
                strSQLcommand = "qryGetAllMaintainTables";
            }
            else
            {
                strSQLcommand = "qryGetStndMaintainTables";
            }

            try
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                SqlCommand cmd = new SqlCommand(strSQLcommand, con);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds, "Tables");
                cboTableName.ItemsSource = ds.Tables[0].DefaultView;
                cboTableName.DisplayMemberPath = ds.Tables[0].Columns["TableName"].ToString();
                cboTableName.SelectedIndex = 0;
                dgQueries.ItemsSource = ds.Tables[0].DefaultView;

                //get Additional_Data_Instruction corresponding to Friendly Name
                strAddlDataInstr = (dgQueries.Items[0] as DataRowView).Row.ItemArray[1].ToString();

                //get Main_Key corresponding to Friendly Name
                strMainKey = (dgQueries.Items[0] as DataRowView).Row.ItemArray[2].ToString();

                FillTablesGrid(strAddlDataInstr);
        }
            catch
            {
                MessageBox.Show("Unable to connect to database.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
}

        private void cboTableName_DropDownClosed(object sender, EventArgs e)
        {
            int intTableIndex = cboTableName.SelectedIndex;
            int intPrevTableIndex = Convert.ToInt16(cboTableName.Tag);
            // where selection unchanged, do nothing 
            if (intTableIndex == intPrevTableIndex)
            {
                return;
            }

            try
            {
                //get Additional_Data_Instruction corresponding to Friendly Name
                strAddlDataInstr = (dgQueries.Items[intTableIndex] as DataRowView).Row.ItemArray[1].ToString();

                //get Main_Key corresponding to Friendly Name
                strMainKey = (dgQueries.Items[intTableIndex] as DataRowView).Row.ItemArray[2].ToString();

                FillTablesGrid(strAddlDataInstr);
            }
            catch
            {
                MessageBox.Show("Select the correct value from the drop down list", "Wealthy Risk Tool", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            


        }

        private void FillTablesGrid(string strAddlDataInstr)
        {
            dgTables.Columns.Clear();

            switch (strAddlDataInstr)
            {
                case "SELECT * FROM [tblAdditional_Data_Sources]":
                    tblAdditional.dt.Clear();
                    tblAdditional.sda.Fill(tblAdditional.dt);
                    dgTables.ItemsSource = tblAdditional.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblAgent_Details] Order By [UTR]":
                    tblAgent.dt.Clear();
                    tblAgent.sda.Fill(tblAgent.dt);
                    dgTables.ItemsSource = tblAgent.dt.DefaultView;
                    break;

                case "SELECT * From [tblInfo]":
                    tblInfo.dt.Clear();
                    tblInfo.sda.Fill(tblInfo.dt);
                    dgTables.ItemsSource = tblInfo.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblAssociation_Types]":
                    tblAssoc.dt.Clear();
                    tblAssoc.sda.Fill(tblAssoc.dt);
                    dgTables.ItemsSource = tblAssoc.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblCombos]":
                    tblCombos.dt.Clear();
                    tblCombos.sda.Fill(tblCombos.dt);
                    dgTables.ItemsSource = tblCombos.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblCRM_Weighting]":
                    tblCRM_W.dt.Clear();
                    tblCRM_W.sda.Fill(tblCRM_W.dt);
                    dgTables.ItemsSource = tblCRM_W.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblCustomer_Data] Order By [UTR]":
                    tblCust.dt.Clear();
                    tblCust.sda.Fill(tblCust.dt);
                    dgTables.ItemsSource = tblCust.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblEmailAddress] Order By [UTR]":
                    tblEmail.dt.Clear();
                    tblEmail.sda.Fill(tblEmail.dt);
                    dgTables.ItemsSource = tblEmail.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblEnquiry_Data] Order By [UTR]":
                    tblEnq.dt.Clear();
                    tblEnq.sda.Fill(tblEnq.dt);
                    dgTables.ItemsSource = tblEnq.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblOfficeCRM]":
                    tblOffice.dt.Clear();
                    tblOffice.sda.Fill(tblOffice.dt);
                    dgTables.ItemsSource = tblOffice.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblQuestion_Weighting] Order By [Control_Name], [Score_entered]":
                    tblQW.dt.Clear();
                    tblQW.sda.Fill(tblQW.dt);
                    dgTables.ItemsSource = tblQW.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblFieldNames]":
                    tblField.dt.Clear();
                    tblField.sda.Fill(tblField.dt);
                    dgTables.ItemsSource = tblField.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblRPD_Calculations]":
                    tblCalc.dt.Clear();
                    tblCalc.sda.Fill(tblCalc.dt);
                    dgTables.ItemsSource = tblCalc.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblRPD_Score_Data] Order By [UTR]":
                    tblScore.dt.Clear();
                    tblScore.sda.Fill(tblScore.dt);
                    dgTables.ItemsSource = tblScore.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblSensitive_Cases_List] Order By [Authorised_User_PID]":
                    tblSensPID.dt.Clear();
                    tblSensPID.sda.Fill(tblSensPID.dt);
                    dgTables.ItemsSource = tblSensPID.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblSensitive_Cases_List] Order By [UTR]":
                    tblSensUTR.dt.Clear();
                    tblSensUTR.sda.Fill(tblSensUTR.dt);
                    dgTables.ItemsSource = tblSensUTR.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblStndReports]":
                    tblRep.dt.Clear();
                    tblRep.sda.Fill(tblRep.dt);
                    dgTables.ItemsSource = tblRep.dt.DefaultView;
                    break;

                case "SELECT * FROM [tblUsers]":
                    tblUser.dt.Clear();
                    tblUser.sda.Fill(tblUser.dt);
                    dgTables.ItemsSource = tblUser.dt.DefaultView;
                    break;

                case "SELECT * FROM [dbo].[ltUserPopID] ORDER BY UserPID, UserPopID":
                    tblUserPopID.dt.Clear();
                    tblUserPopID.sda.Fill(tblUserPopID.dt);
                    dgTables.ItemsSource = tblUserPopID.dt.DefaultView;

                    // define DataGridTemplateColumn 
                    DataGridTemplateColumn dgTemplateColumn = new DataGridTemplateColumn();
                    dgTemplateColumn.Header = "Populations";
                    dgTables.Columns[2].Visibility = Visibility.Hidden; // hide corresponding standard column

                    // define ComboBox
                    var cbo = new FrameworkElementFactory(typeof(ComboBox));
                    cbo.SetValue(ComboBox.NameProperty, "cboPopulations");
                    SqlConnection conn = new SqlConnection(Global.ConnectionString);
                    SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM dbo.tblPopulations", conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "dbo.tblPopulations");
                    Binding bdCombo = new Binding("");
                    bdCombo.Mode = BindingMode.OneWay;
                    bdCombo.Source = ds.Tables[0].DefaultView;
                    cbo.SetBinding(ComboBox.ItemsSourceProperty, bdCombo);
                    cbo.SetValue(ComboBox.SelectedValuePathProperty, "PopID");
                    cbo.SetValue(ComboBox.DisplayMemberPathProperty, "Pop_Friendly_Name");

                    // bind ComboBox to DataGrid
                    var bdValue = new Binding("UserPopID")
                    {
                        Mode = BindingMode.TwoWay,
                        UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                    };
                    cbo.SetBinding(ComboBox.SelectedValueProperty, bdValue);
                    var dt = new DataTemplate();
                    dt.VisualTree = cbo;
                    dgTemplateColumn.CellTemplate = dt;

                    // add combo to the DataGrid
                    dgTables.Columns.Add(dgTemplateColumn);
                    break;
            }


        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            try
            {


                switch (strAddlDataInstr)
                {
                    case "SELECT * FROM [tblAdditional_Data_Sources]":
                        tblAdditional.conn.Open();
                        SqlCommandBuilder bldAdd = new SqlCommandBuilder(tblAdditional.sda);
                        tblAdditional.sda.UpdateCommand = bldAdd.GetUpdateCommand();
                        tblAdditional.sda.Update(tblAdditional.dt);
                        tblAdditional.conn.Close();
                        break;

                    case "SELECT * FROM [tblAgent_Details] Order By [UTR]":
                        tblAgent.conn.Open();
                        SqlCommandBuilder bldAgent = new SqlCommandBuilder(tblAgent.sda);
                        tblAgent.sda.UpdateCommand = bldAgent.GetUpdateCommand();
                        tblAgent.sda.Update(tblAgent.dt);
                        tblAgent.conn.Close();
                        break;

                    case "SELECT * From [tblInfo]":
                        tblInfo.conn.Open();
                        SqlCommandBuilder bldInfo = new SqlCommandBuilder(tblInfo.sda);
                        tblInfo.sda.UpdateCommand = bldInfo.GetUpdateCommand();
                        tblInfo.sda.Update(tblInfo.dt);
                        tblInfo.conn.Close();
                        break;

                    case "SELECT * FROM [tblAssociation_Types]":
                        tblAssoc.conn.Open();
                        SqlCommandBuilder bldAssoc = new SqlCommandBuilder(tblAssoc.sda);
                        tblAssoc.sda.UpdateCommand = bldAssoc.GetUpdateCommand();
                        tblAssoc.sda.Update(tblAssoc.dt);
                        tblAssoc.conn.Close();
                        break;

                    case "SELECT * FROM [tblCombos]":
                        tblCombos.conn.Open();
                        SqlCommandBuilder bldCombos = new SqlCommandBuilder(tblCombos.sda);
                        tblCombos.sda.UpdateCommand = bldCombos.GetUpdateCommand();
                        tblCombos.sda.Update(tblCombos.dt);
                        tblCombos.conn.Close();
                        break;

                    case "SELECT * FROM [tblCRM_Weighting]":
                        tblCRM_W.conn.Open();
                        SqlCommandBuilder bldCRM_W = new SqlCommandBuilder(tblCRM_W.sda);
                        tblCRM_W.sda.UpdateCommand = bldCRM_W.GetUpdateCommand();
                        tblCRM_W.sda.Update(tblCRM_W.dt);
                        tblCRM_W.conn.Close();
                        break;

                    case "SELECT * FROM [tblCustomer_Data] Order By [UTR]":
                        tblCust.conn.Open();
                        SqlCommandBuilder bldCust = new SqlCommandBuilder(tblCust.sda);
                        tblCust.sda.UpdateCommand = bldCust.GetUpdateCommand();
                        tblCust.sda.Update(tblCust.dt);
                        tblCust.conn.Close();
                        break;

                    case "SELECT * FROM [tblEmailAddress] Order By [UTR]":
                        tblEmail.conn.Open();
                        SqlCommandBuilder bldEmail = new SqlCommandBuilder(tblEmail.sda);
                        tblEmail.sda.UpdateCommand = bldEmail.GetUpdateCommand();
                        tblEmail.sda.Update(tblEmail.dt);
                        tblEmail.conn.Close();
                        break;

                    case "SELECT * FROM [tblEnquiry_Data] Order By [UTR]":
                        tblEnq.conn.Open();
                        SqlCommandBuilder bldEnq = new SqlCommandBuilder(tblEnq.sda);
                        tblEnq.sda.UpdateCommand = bldEnq.GetUpdateCommand();
                        tblEnq.sda.Update(tblEnq.dt);
                        tblEnq.conn.Close();
                        break;

                    case "SELECT * FROM [tblOfficeCRM]":
                        tblOffice.conn.Open();
                        SqlCommandBuilder bldOffice = new SqlCommandBuilder(tblOffice.sda);
                        tblOffice.sda.UpdateCommand = bldOffice.GetUpdateCommand();
                        tblOffice.sda.Update(tblOffice.dt);
                        tblOffice.conn.Close();
                        break;

                    case "SELECT * FROM [tblQuestion_Weighting] Order By [Control_Name], [Score_entered]":
                        tblQW.conn.Open();
                        SqlCommandBuilder bldQW = new SqlCommandBuilder(tblQW.sda);
                        tblQW.sda.UpdateCommand = bldQW.GetUpdateCommand();
                        tblQW.sda.Update(tblQW.dt);
                        tblQW.conn.Close();
                        break;

                    case "SELECT * FROM [tblFieldNames]":
                        tblField.conn.Open();
                        SqlCommandBuilder bldField = new SqlCommandBuilder(tblField.sda);
                        tblField.sda.UpdateCommand = bldField.GetUpdateCommand();
                        tblField.sda.Update(tblField.dt);
                        tblField.conn.Close();
                        break;

                    case "SELECT * FROM [tblRPD_Calculations]":
                        tblCalc.conn.Open();
                        SqlCommandBuilder bldCalc = new SqlCommandBuilder(tblCalc.sda);
                        tblCalc.sda.UpdateCommand = bldCalc.GetUpdateCommand();
                        tblCalc.sda.Update(tblCalc.dt);
                        tblCalc.conn.Close();
                        break;

                    case "SELECT * FROM [tblRPD_Score_Data] Order By [UTR]":
                        tblScore.conn.Open();
                        SqlCommandBuilder bldScore = new SqlCommandBuilder(tblScore.sda);
                        tblScore.sda.UpdateCommand = bldScore.GetUpdateCommand();
                        tblScore.sda.Update(tblScore.dt);
                        tblScore.conn.Close();
                        break;

                    case "SELECT * FROM [tblSensitive_Cases_List] Order By [Authorised_User_PID]":
                        tblSensPID.conn.Open();
                        SqlCommandBuilder bldSPID = new SqlCommandBuilder(tblSensPID.sda);
                        tblSensPID.sda.UpdateCommand = bldSPID.GetUpdateCommand();
                        tblSensPID.sda.Update(tblSensPID.dt);
                        tblSensPID.conn.Close();
                        break;

                    case "SELECT * FROM [tblSensitive_Cases_List] Order By [UTR]":
                        tblSensUTR.conn.Open();
                        SqlCommandBuilder bldSUTR = new SqlCommandBuilder(tblSensUTR.sda);
                        tblSensUTR.sda.UpdateCommand = bldSUTR.GetUpdateCommand();
                        tblSensUTR.sda.Update(tblSensUTR.dt);
                        tblSensUTR.conn.Close();
                        break;

                    case "SELECT * FROM [tblstndReports]":
                        tblRep.conn.Open();
                        SqlCommandBuilder bldRep = new SqlCommandBuilder(tblRep.sda);
                        tblRep.sda.UpdateCommand = bldRep.GetUpdateCommand();
                        tblRep.sda.Update(tblRep.dt);
                        tblRep.conn.Close();
                        break;

                    case "SELECT * FROM [tblUsers]":
                        tblUser.conn.Open();
                        SqlCommandBuilder bldUser = new SqlCommandBuilder(tblUser.sda);
                        tblUser.sda.UpdateCommand = bldUser.GetUpdateCommand();
                        tblUser.sda.Update(tblUser.dt);
                        tblUser.conn.Close();
                        break;

                    case "SELECT * FROM [dbo].[ltUserPopID] ORDER BY UserPID, UserPopID":
                        tblUserPopID.conn.Open();
                        SqlCommandBuilder bldUserPop = new SqlCommandBuilder(tblUserPopID.sda);
                        tblUserPopID.sda.UpdateCommand = bldUserPop.GetUpdateCommand();
                        tblUserPopID.sda.Update(tblUserPopID.dt);
                        tblUserPopID.conn.Close();
                        break;


                }

            }
            catch(Exception ex)
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                MessageBox.Show(ex.ToString(),"Wealthy Risk Tool",MessageBoxButton.OK, MessageBoxImage.Error);
            }
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }


        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CboTableName_DropDownOpened(object sender, EventArgs e)
        {
            cboTableName.Tag = cboTableName.SelectedIndex;
        }
    }
}
