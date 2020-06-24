using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for RPT_Detail.xaml
    /// </summary>
    public partial class RPT_Detail : Window
    {
        TextBoxDate tbd = new TextBoxDate();
        bool bFormLoaded = false;
        public int WeightedScore { get; set; }

        public RPT_Detail()
        {
            this.DataContext = new RPT(); // set data context
            InitializeComponent();
            PopulateCombos();

            this.WindowState = WindowState.Maximized;

        }


        private void PopulateCombos()
        {
            Dictionary<int, string> yesnodictionary = new Dictionary<int, string>();
            yesnodictionary.Add(0, "No");
            yesnodictionary.Add(1, "Yes");

            try
            {
                Lookups lu = new Lookups();
                lu.GetRPTDetailLookups();
                lu.GetRPTDetailOfficeCRMs();
                //lu.GetOfficeCRMs();

                // == Customer Segment

                // cboSegment
                cboStrand.ItemsSource = lu.dsRPTDetailCombo.Tables[13].DefaultView;
                cboStrand.DisplayMemberPath = "Options";
                cboStrand.SelectedValuePath = "DecodedValue";

                // cboSegment
                cboSegment.ItemsSource = lu.dsRPTDetailCombo.Tables[0].DefaultView;
                cboSegment.DisplayMemberPath = "Options";
                cboSegment.SelectedValuePath = "DecodedValue";

                // cboPopFriendly
                cboPopFriendly.ItemsSource = lu.dsRPTDetailCombo.Tables[1].DefaultView;
                cboPopFriendly.DisplayMemberPath = "Pop_Friendly_Name";
                cboPopFriendly.SelectedValuePath = "Pop_Code_Name";

                // cboPopCode
                cboPopCode.ItemsSource = lu.dsRPTDetailCombo.Tables[1].DefaultView;
                cboPopCode.DisplayMemberPath = "Pop_Friendly_Name";
                cboPopCode.SelectedValuePath = "Pop_Code_Name";

                // == Customer Details

                // cboUTR
                //cboUTR.ItemsSource = lu.dsRPTDetailCombo.Tables[2].DefaultView;
                //cboUTR.DisplayMemberPath = "Options";
                //cboUTR.SelectedValuePath = "DecodedValue";

                // cboDeceased [yes/no]
                cboDeceased.ItemsSource = new System.Windows.Forms.BindingSource(yesnodictionary, null);
                cboDeceased.DisplayMemberPath = "Value";
                cboDeceased.SelectedValuePath = "Key";

                // cboMarital
                cboMarital.ItemsSource = lu.dsRPTDetailCombo.Tables[2].DefaultView;
                cboMarital.DisplayMemberPath = "Options";
                cboMarital.SelectedValuePath = "DecodedValue";

                // cboGender
                cboGender.ItemsSource = lu.dsRPTDetailCombo.Tables[3].DefaultView;
                cboGender.DisplayMemberPath = "Options";
                cboGender.SelectedValuePath = "DecodedValue";

                // cboResidence
                cboResidence.ItemsSource = lu.dsRPTDetailCombo.Tables[4].DefaultView;
                cboResidence.DisplayMemberPath = "Options";
                cboResidence.SelectedValuePath = "DecodedValue";

                // cboDomicile
                cboDomicile.ItemsSource = lu.dsRPTDetailCombo.Tables[5].DefaultView;
                cboDomicile.DisplayMemberPath = "Options";
                cboDomicile.SelectedValuePath = "DecodedValue";

                // == Customer Team

                // cboOffice populated on Window_Loaded() / PopulateAndSetAllocationCombos() in order to pass selected PopCode
                // cboTeam  populated on Window_Loaded() / PopulateAndSetAllocationCombos()  in order to pass selected Office and PopCode
                // cboAllocatedTo: populated on Window_Loaded() / PopulateAndSetAllocationCombos()  in order to pass selected Office and Team

                //cboCRMName
                cboCRMName.ItemsSource = lu.dsRPTDetailOfficeCRMs.Tables[0].DefaultView;
                cboCRMName.DisplayMemberPath = "CRM_Name";
                cboCRMName.SelectedValuePath = "CRM_Name";

                // == Appointed Agent

                //cbo648 [yes/no]
                cbo648.ItemsSource = new System.Windows.Forms.BindingSource(yesnodictionary, null);
                cbo648.DisplayMemberPath = "Value";
                cbo648.SelectedValuePath = "Key";

                //cboChange [yes/no]
                cboChange.ItemsSource = new System.Windows.Forms.BindingSource(yesnodictionary, null);
                cboChange.DisplayMemberPath = "Value";
                cboChange.SelectedValuePath = "Key";

                // == Behaviors

                //cboCurrentSuspensions [yes/no/unknown]
                cboCurrentSuspensions.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboCurrentSuspensions.DisplayMemberPath = "Options";
                cboCurrentSuspensions.SelectedValuePath = "DecodedValue";

                //cboPrevSuspensions [yes/no/unknown]
                cboPrevSuspensions.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboPrevSuspensions.DisplayMemberPath = "Options";
                cboPrevSuspensions.SelectedValuePath = "DecodedValue";

                //cboFailures [yes/no/unknown]
                cboFailures.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboFailures.DisplayMemberPath = "Options";
                cboFailures.SelectedValuePath = "DecodedValue";

                // == Case overview tab
                //cboWealth
                cboWealth.ItemsSource = lu.dsRPTDetailCombo.Tables[6].DefaultView;
                cboWealth.DisplayMemberPath = "Options";
                cboWealth.SelectedValuePath = "DecodedValue";
                //cboSector
                cboSector.ItemsSource = lu.dsRPTDetailCombo.Tables[9].DefaultView;
                cboSector.DisplayMemberPath = "Options";
                cboSector.SelectedValuePath = "DecodedValue";
                //cboPathway
                cboPathway.ItemsSource = lu.dsRPTDetailCombo.Tables[7].DefaultView;
                cboPathway.DisplayMemberPath = "Options";
                cboPathway.SelectedValuePath = "DecodedValue";
                //cboLongTerm
                cboLongTerm.ItemsSource = lu.dsRPTDetailCombo.Tables[10].DefaultView;
                cboLongTerm.DisplayMemberPath = "Options";
                cboLongTerm.SelectedValuePath = "DecodedValue";
                //cboSource
                cboSource.ItemsSource = lu.dsRPTDetailCombo.Tables[8].DefaultView;
                cboSource.DisplayMemberPath = "Options";
                cboSource.SelectedValuePath = "DecodedValue";
                //cboLifeEvents
                cboLifeEvents.ItemsSource = lu.dsRPTDetailCombo.Tables[11].DefaultView;
                cboLifeEvents.DisplayMemberPath = "Options";
                cboLifeEvents.SelectedValuePath = "DecodedValue";
            }
            catch
            {
                System.Windows.MessageBox.Show("Unable to load combo box data.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            PopulateAndSetAllocationCombos();

            if (Globals.gs_CRM.Count < 1) /*test for weightings*/
            {
                RPT.RPT_Data rpt = new RPT.RPT_Data();
                rpt.GetCRM();
            }

            int iTest = Convert.ToInt16(Globals.gs_CRM.ElementAt(0)[1]);

            //TabFrames
            //this.Activated += AfterLoading;

            bFormLoaded = true;
        }


        private void PopulateAndSetAllocationCombos()
        {
            bool blnTest = PopulateOfficeCombo();
            blnTest = PopulateTeamCombo();
            blnTest = PopulateAllocatedToCombo();
        }

        private bool PopulateOfficeCombo()
        {
            string strPopCode = "";
            Lookups lu = new Lookups();
            try
            {
                // populate cboOffice, based on PopCode
                strPopCode = this.cboPopCode.SelectedValue.ToString();
                lu.GetOffices(strPopCode);
                cboOffice.ItemsSource = lu.dsOffices.Tables[0].DefaultView;
                cboOffice.DisplayMemberPath = "Office";
                cboOffice.SelectedValuePath = "Office";
                return true;
            }
            catch
            {
                return false;
            };
        }

        private bool PopulateTeamCombo()
        {
            string strPopCode = "";
            string strOffice = "";
            Lookups lu = new Lookups();
            try
            {
                // populate cboTeam, based on Office and PopCode
                strOffice = this.cboOffice.SelectedValue.ToString();
                strPopCode = this.cboPopCode.SelectedValue.ToString();
                lu.GetOfficeCRMs(strOffice, strPopCode); // "rPt20Mill"
                cboTeam.ItemsSource = lu.dsOfficeTeams.Tables[0].DefaultView;
                cboTeam.DisplayMemberPath = "Team Identifier";
                cboTeam.SelectedValuePath = "Team Identifier";
                return true;
            }
            catch
            {
                return false;
            };
        }

        private bool PopulateAllocatedToCombo()
        {
            string strOffice = "";
            string strTeam = "";
            Lookups lu = new Lookups();
            try
            {
                // populate cboAllocatedTo, based on Office and Team
                strOffice = this.cboOffice.SelectedValue.ToString();
                strTeam = this.cboTeam.SelectedValue.ToString();
                lu.GetOfficeTeamStaff(strOffice, strTeam);
                cboAllocatedTo.ItemsSource = lu.dsOfficeTeamStaff.Tables[0].DefaultView;
                cboAllocatedTo.DisplayMemberPath = "AllocatedTo";
                cboAllocatedTo.SelectedValuePath = "AllocatedTo";
                return true;
            }
            catch
            {
                return false;
            };
        }

        private void TxtDOB_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = tbd.TextBox_KeyDown(e);
        }

        private void TxtDOB_LostFocus(object sender, RoutedEventArgs e)
        {

            e.Handled = !tbd.TextBox_LostFocus(txtDOB);

        }

        private void TxtDOD_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = tbd.TextBox_KeyDown(e);
        }

        private void TxtDOD_LostFocus(object sender, RoutedEventArgs e)
        {
            e.Handled = !tbd.TextBox_LostFocus(txtDOD);
        }

        #region GotFocus ShowActiveControl

        private void ShowActiveControl(System.Windows.Controls.Control ctl)
        {
            string strName = ctl.Name.Substring(3, ctl.Name.Length - 3); // remove 3 character prefix e.g. cbo, cmd, lvw
            ((TabItem)tabRPD.Items[0]).Header = "Customer";
            ((TabItem)tabRPD.Items[1]).Header = "Case Overview";
            ((TabItem)tabRPD.Items[2]).Header = "Information";
            ((TabItem)tabRPD.Items[3]).Header = "Results";
            switch (tabRPD.SelectedIndex)
            {
                case 0:
                    strName = "Customer - " + strName;
                    break;
                case 1:
                    strName = "Case Overview - " + strName;
                    break;
                case 2:
                    strName = "Information - " + strName;
                    break;
                case 3:
                    strName = "Results - " + strName;
                    break;
                default:
                    break;
            }
            ((TabItem)tabRPD.Items[tabRPD.SelectedIndex]).Header = strName;
        }

        private void CboStrand_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboStrand);
            }
        }

        private void TxtSurname_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtSurname);
            }
        }

        private void TxtFirstName_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtFirstName);
            }
        }

        private void TxtUTR_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtUTR);
            }
        }

        private void CboDeceased_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboDeceased);
            }
        }

        private void TxtDOB_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtDOB);
            }
        }

        private void TxtDOD_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtDOD);
            }
        }

        private void CboMarital_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboMarital);
            }
        }

        private void CboGender_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboGender);
            }
        }

        private void TxtPrivateAddress_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtPrivateAddress);
            }
        }

        private void TxtPostcode_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtPostcode);
            }
        }

        private void TxtDeselected_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtDeselected);
            }
        }

        private void ChkDeselected_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(chkDeselected);
            }
        }

        private void TxtSecondaryAddress_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtSecondaryAddress);
            }
        }

        private void CboResidence_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboResidence);
            }
        }

        private void CboDomicile_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboDomicile);
            }
        }

        private void CboOffice_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboOffice);
            }
        }

        private void CboTeam_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboTeam);
            }
        }

        private void CboAllocatedTo_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboAllocatedTo);
            }
        }

        private void CboCRMName_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboCRMName);
            }
        }

        private void TxtCRMDA_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtCRMDA);
            }
        }

        private void TxtFirm_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtFirm);
            }
        }

        private void TxtTelNo_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtTelNo);
            }
        }

        private void TxtAgentCode_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtAgentCode);
            }
        }

        private void TxtAgentAddress_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtAgentAddress);
            }
        }

        private void TxtNamedAgent_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtNamedAgent);
            }
        }


        private void TxtOtherContact_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtOtherContact);
            }
        }

        private void Cbo648_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cbo648);
            }
        }

        private void CboChange_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboChange);
            }
        }

        private void CboWealth_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboWealth);
            }
        }

        private void CboSector_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboSector);
            }
        }

        private void CboPathway_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboPathway);
            }
        }

        private void CboLongTerm_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboLongTerm);
            }
        }

        private void CboSource_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboSource);
            }
        }

        private void CboLifeEvents_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboLifeEvents);
            }
        }

        private void TxtNarrative_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtNarrative);
            }
        }

        private void LvwAssociates_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(lvwAssociates);
            }
        }

        private void lvwEmail_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(lvwEmail);
            }
        }

        private void TxtOpenRisks_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtOpenRisks);
            }
        }

        private void TxtSettled_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtSettled);
            }
        }

        private void TxtHighestSettled_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtHighestSettled);
            }
        }

        private void TxtHighestPercent_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtHighestPercent);
            }
        }

        private void CboCurrentSuspensions_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboCurrentSuspensions);
            }
        }

        private void CboPrevSuspensions_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboPrevSuspensions);
            }
        }

        private void CboFailures_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboFailures);
            }
        }

        private void TxtOpenIDMS_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtOpenIDMS);
            }
        }

        private void TxtClosedIDMS_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtClosedIDMS);
            }
        }

        private void CboSegment_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboSegment);
            }
        }

        private void CboPopFriendly_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cboPopFriendly);
            }
        }

        private void CmdEmailAdd_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdEmailAdd);
            }
        }

        private void CmdEmailUpdate_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdEmailUpdate);
            }
        }

        private void CmdEmailDelete_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdEmailDelete);
            }
        }

        private void CmdAddAssociate_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdAddAssociate);
            }
        }

        private void CmdUpdateAssociate_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdUpdateAssociate);
            }
        }

        private void CmdDeleteAssociate_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdDeleteAssociate);
            }
        }

        private void LvwAdditionalInfo_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(lvwAdditionalInfo);
            }
        }

        private void RtbAdditionalData_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(rtbAdditionalData);
            }
        }

        private void CmdAddAction_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdAddAction);
            }
        }

        //private void LvwPrevRes_GotFocus(object sender, RoutedEventArgs e)
        //{
        //    if (Globals.blnAccess == true)
        //    {
        //        ShowActiveControl(lvwPrevRes);
        //    }
        //}

        private void MscHistory_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(mscHistory);
            }
        }

        private void CmdGuidance_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdGuidance);
            }
        }

        private void CmdSave_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdSave);
            }
        }

        private void CmdUpdateClose_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdUpdateClose);
            }
        }

        private void CmdCancel_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(cmdCancel);
            }
        }

        #endregion

        private void CmdCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CboPopFriendly_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((cboPopCode.SelectedIndex != cboPopFriendly.SelectedIndex)&&(bFormLoaded == true))
            {
                if (System.Windows.Forms.MessageBox.Show
                            ("You are about to change the Records Population and you will also need to update the Office, Team and Allocation."
                    + Environment.NewLine + Environment.NewLine
                    + "You may also need to consider changing the CCM."
                    + Environment.NewLine + Environment.NewLine
                    + "Select Yes to Proceed", "Population", System.Windows.Forms.MessageBoxButtons.YesNo,
                            System.Windows.Forms.MessageBoxIcon.Exclamation, System.Windows.Forms.MessageBoxDefaultButton.Button2)
                            == System.Windows.Forms.DialogResult.No)
                {
                    // aborted so reset population
                    cboPopFriendly.SelectedIndex = cboPopCode.SelectedIndex;
                }
                else
                {
                    // user confirmed yes so change population
                    cboPopCode.SelectedIndex = cboPopFriendly.SelectedIndex;
                    txtPopFriendly.Text = cboPopFriendly.Text;
                    // update Office options as not all populations may have same office locations
                    // ##########cboOffice.Items.Clear();
                    bool blnTest = PopulateOfficeCombo();
                    foreach (var item in cboOffice.Items)
                    {
                        if (item.ToString() == txtOffice.Text)
                        {
                            cboOffice.Text = txtOffice.Text;
                            break;
                        }
                    }
                    // ###########cboTeam.Items.Clear();
                }
            }
        }


        private void ChkDeselected_Checked(object sender, RoutedEventArgs e)
        {
            if (System.Windows.Forms.MessageBox.Show
                        ("Please confirm that you wish to deselect this case.", "Deselection", System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Exclamation, System.Windows.Forms.MessageBoxDefaultButton.Button2)
                        == System.Windows.Forms.DialogResult.No)
            {
                chkDeselected.IsChecked = false;
            }
            else
            {
                txtDeselected.Text = DateTime.Now.ToString("dd/MM/yyyy");
            }

        }

        private void ChkDeselected_Unchecked(object sender, RoutedEventArgs e)
        {
            txtDeselected.Tag = txtDeselected.Text; /*store for possible recall*/
            txtDeselected.Text = "";
        }

        private void ChkCRMDescretion_Checked(object sender, RoutedEventArgs e)
        {
           if (Globals.gs_CRM.Count < 1) /*test for weightings*/
            {
                RPT.RPT_Data rpt = new RPT.RPT_Data();
                rpt.GetCRM();
            }
            int iWeighting = CRMWeighting(cboPopCode.SelectedValue.ToString().ToUpper());
            txtCRMScore.Text = iWeighting.ToString();
            txtCRMExplanation.Text = "";
            txtCRMExplanation.IsEnabled = true;
            RecalculateBehaviours();
            //RecalculateResults();
        }

        private int CRMWeighting(string sPop)
        {
            // return weighting, for given population
            int iCRMWeighting = 0;
            if (Globals.gs_CRM.Count > 0)
            {
                foreach (var crm in Globals.gs_CRM)
                {
                    if (crm.ElementAt(0).ToUpper() == sPop)
                    {
                        iCRMWeighting = Convert.ToInt16(crm.ElementAt(1));
                    }
                }
            }
            return iCRMWeighting;
        }

        private void ChkCRMDescretion_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCRMScore.Text = "";
            txtCRMExplanation.Text = "";
            txtCRMExplanation.IsEnabled = false;
            RecalculateBehaviours();
            //RecalculateResults();
        }

        private void CmdSave_Click(object sender, RoutedEventArgs e)
        {
            RecalculateBehaviours();

            //BindingExpression obj = txtSecondaryAddress.GetBindingExpression(TextBox.TextProperty);
            //obj.UpdateSource();
            SaveRecord();
        }

        private void SaveRecord()
        {
            this.cmdSave.IsEnabled = false;
            Globals.blnUpdate = false;

            RPT.RPT_Data rpt = new RPT.RPT_Data(); // initialise data

            if (txtHighestPercent.Text != "" && IsNumeric(txtHighestPercent.Text) == false)
            {
                if (Convert.ToInt16(txtHighestPercent.Text) > 500)
                {
                    System.Windows.MessageBox.Show("The highest percentage penalty value is too high.  Please check and re-enter.", "Highest Percentage", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtHighestPercent.Text = rpt.Highest_Percentage.ToString();
                    cmdSave.IsEnabled = true;
                    txtHighestPercent.Focus();
                    return;
                }
            }

            if ((cboCRMName.Text.Trim() == "") || (txtCRMDA.Text.Trim() == ""))
            {
                System.Windows.MessageBox.Show("Please confirm the CRM for this case" + Environment.NewLine + "and/or the date of their appointment.", "CRM", MessageBoxButton.OK, MessageBoxImage.Information);
                cmdSave.IsEnabled = true;
                return;
            }

            // run through each element to check if it needs updating and then update relevant elements.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            CheckQuestions(); // ensure questions are up to date with weights

            //RecalculateResults();  // ensure scores are up to date

            //if (tcmdSave.Text.IndexOf("save", StringComparison.CurrentCultureIgnoreCase) >= 0)
            //{
                if ((cboOffice.Text.Trim() == "") || (cboTeam.Text.Trim() == "") || (cboPopFriendly.Text.Trim() == ""))
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                    MessageBox.Show("Please assign a Population, an Office and a Team.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                    cmdSave.IsEnabled = true;
                    return;
                }

                if ((txtCRMExplanation.Text.Trim() == "") && (chkCRMDescretion.IsChecked == true))
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                    MessageBox.Show("Please give an explanation of why CCM Discretion has been applied.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                    cmdSave.IsEnabled = true;
                    return;
                }

                if (CheckandAddRPD() == false)
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                    MessageBox.Show("Problem adding data. Please try again." + Environment.NewLine + Environment.NewLine + "If the problem persists, take screenshots of the" + Environment.NewLine + "record and error message, and report it to ICT.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                    cmdSave.IsEnabled = true;
                }
                else
                {
                    // reload new updated cRPD Data
                    GetRPDDataFromForm();

                    double dUTR = 0;
                    try { dUTR = double.Parse(txtUTR.Text); } catch { dUTR = 0; };
                    int iYear = 2000;
                    try { iYear = Convert.ToInt16(lblPopYear.Text); } catch { }
                    double dPercentile = 0;
                    try { dPercentile = Convert.ToDouble(txtPercentile.Text.ToString()); } catch { dPercentile = 0; } // text can be 'N/A'
                    string sPop = "";
                    try { sPop = cboPopCode.SelectedValue.ToString(); } catch { }
                    if (rpt.GetRPDHistoricalData(dUTR, iYear, dPercentile, sPop) == false)
                    {
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        MessageBox.Show("Data Saved but issue repopulating previous score data.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        MessageBox.Show("Data saved.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }

            //}
            //else  // i.e. if(tcmdSave.Text != Save)
            //{
            //    MessageBox.Show("There has been an error loading this form." + Environment.NewLine + "Please notify the Support Team as this will need debugging.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
            //}

            cmdSave.IsEnabled = true;
            Globals.blnUpdate = true;

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

            private void CmdUTR_Click(object sender, RoutedEventArgs e)
        {
            string sUTR = txtUTR.Text.ToString().Trim();
            if ((sUTR == "") || (sUTR.Length != 10)) /*UTR is blank or not 10 characters long*/
            {
                return; /*do nothing*/
            }
            else
            {
                /*copy UTR to Clipboard*/
                Clipboard.Clear();
                Clipboard.SetText(sUTR);
            }
        }

        private void CboOffice_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if((bFormLoaded == true))
            {
                bool blnTest = PopulateTeamCombo();
                cboTeam.Text = "";
                blnTest = PopulateAllocatedToCombo();
                cboAllocatedTo.Text = "";
                Globals.blnOfficeWiped = true;
                return;
            }
            else
            {
                if (Globals.blnOfficeWiped == true)
                {
                    Globals.blnOfficeWiped = false;
                }
                else
                {
                    if (txtOffice.Text == cboOffice.Text)
                    {
                        return;
                    }
                }
            }

            if (cboPopFriendly.Text == "")
            {
                MessageBox.Show("Please confirm the Record Population first.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                cboPopFriendly.Focus();
                cboPopFriendly.IsDropDownOpen = true;
                return;
            }

            //PopulateCombo Me.cboTeam, "qryGetOfficeTeams", "Team Identifier", Me.cboOffice.Text, True, cUser.Pop
            //Me.cboAllocatedTo.Clear
            //Me.cboAllocatedTo.Text = ""

            //Me.txtOffice.Text = Me.cboOffice.Text
            //Me.txtTeam.Text = Me.cboTeam.Text

            //Me.cboTeam.Enabled = True
        }

        private void TxtUTR_LostFocus(object sender, RoutedEventArgs e)
        {
            bool bChecked = false;
            if (txtUTR.Text ==  "")
            {
                return;
            }

            if (txtUTR.Text.Length != 10)
            {
                System.Windows.MessageBox.Show("UTR format appears to be invalid.  Please rectify." + Environment.NewLine + "(10 character numerical string required)", "UTR", MessageBoxButton.OK, MessageBoxImage.Information);
                e.Handled = true;
                return;
            }

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            double dUTR = 0;
            try { dUTR = double.Parse(txtUTR.Text); } catch { dUTR = 0; }; 
            RPT.RPT_Data rpt = new RPT.RPT_Data();
            bChecked =  rpt.CheckCustomer(dUTR);

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            if (bChecked == true)
            {
                System.Windows.MessageBox.Show("A Customer Record already exists for this UTR.", "UTR", MessageBoxButton.OK, MessageBoxImage.Information);
                txtUTR.Text = "";
                return;
            }

            bChecked = rpt.GetCRMMEnquiryDataScore(dUTR);
            if (bChecked == true)
            {
                txtOpenRisks.Text = rpt.Risks_Open.ToString();
                txtSettled.Text = rpt.Settled_Risks.ToString();
                txtHighestSettled.Text = rpt.Highest_Settlement.ToString();
                txtHighestPercent.Text = rpt.Highest_Percentage.ToString();
                if(rpt.CRMM_Date_Added == "")
                {
                    grpCustomer.Header = "Enquiry Data - last updated "; 
                }
                else
                {
                    grpCustomer.Header = "Enquiry Data - last updated " + DateTime.Now.ToString("dd MMM yyyy"); 
                }
            }
            else
            {
                txtOpenRisks.Text = "0";
                txtSettled.Text = "0";
                txtHighestSettled.Text = "0";
                txtHighestPercent.Text = "0";
                grpCustomer.Header = "Enquiry Data - last updated " + DateTime.Now.ToString("dd MMM yyyy");
            }

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

            System.Windows.MessageBox.Show("Error checking UTR for existing customer record.", "UTR", MessageBoxButton.OK, MessageBoxImage.Information);

        }

        private void CheckQuestions()
        {
            txtOpenIDMS.RaiseEvent(new RoutedEventArgs(LostFocusEvent, txtOpenIDMS));
            txtClosedIDMS.RaiseEvent(new RoutedEventArgs(LostFocusEvent, txtClosedIDMS));
            txtOpenRisks.RaiseEvent(new RoutedEventArgs(LostFocusEvent, txtOpenRisks));
            txtSettled.RaiseEvent(new RoutedEventArgs(LostFocusEvent, txtSettled));
            txtHighestPercent.RaiseEvent(new RoutedEventArgs(LostFocusEvent, txtHighestPercent));
            txtHighestSettled.RaiseEvent(new RoutedEventArgs(LostFocusEvent, txtHighestSettled));
            cboCurrentSuspensions.RaiseEvent(new RoutedEventArgs(LostFocusEvent, cboCurrentSuspensions));
            cboPrevSuspensions.RaiseEvent(new RoutedEventArgs(LostFocusEvent, cboPrevSuspensions));
            cboFailures.RaiseEvent(new RoutedEventArgs(LostFocusEvent, cboFailures));
        }

        private void TxtOpenIDMS_LostFocus(object sender, RoutedEventArgs e)
        {
            if((txtOpenIDMS.Text != "")||(IsNumeric(txtOpenIDMS.Text)==true))
            {
                CheckWeight(txtOpenIDMS, txtOpenIDMSHD);
            }
        }

        private void TxtClosedIDMS_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((txtClosedIDMS.Text != "") || (IsNumeric(txtClosedIDMS.Text) == true))
            {
                CheckWeight(txtClosedIDMS, txtClosedIDMSHD);
            }
        }

        private void TxtOpenRisks_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((txtOpenRisks.Text != "") || (IsNumeric(txtOpenRisks.Text) == true))
            {
                CheckWeight(txtOpenRisks, txtOpenRisksHD);
            }
        }

        private void TxtSettled_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((txtSettled.Text != "") || (IsNumeric(txtSettled.Text) == true))
            {
                CheckWeight(txtSettled, txtSettledHD);
            }
        }

        private void TxtHighestPercent_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((txtHighestPercent.Text != "") || (IsNumeric(txtHighestPercent.Text) == true))
            {
                CheckWeight(txtHighestPercent, txtHighestPercentHD);
            }
        }

        private void TxtHighestSettled_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((txtHighestSettled.Text != "") || (IsNumeric(txtHighestSettled.Text) == true))
            {
                CheckWeight(txtHighestSettled, txtHighestSettledHD);
            }
        }

        private void CboCurrentSuspensions_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((cboCurrentSuspensions.Text != "") || (IsNumeric(cboCurrentSuspensions.Text) == true))
            {
                //CheckWeight(txtCurrentSuspensions, txtCurrentSuspensionsHD);
            }
        }

        private void CboPrevSuspensions_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((cboPrevSuspensions.Text != "") || (IsNumeric(cboPrevSuspensions.Text) == true))
            {
                //CheckWeight(txtPrevSuspensions, txtPrevSuspensionsHD);
            }
        }

        private void CboFailures_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((cboFailures.Text != "") || (IsNumeric(cboFailures.Text) == true))
            {
                //CheckWeight(txtFailures, txtFailuresHD);
            }
        }

        public static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public void CheckWeight(TextBox tmpTextBox, TextBox tmpHDTextBox)
        {
            string sWeight = "";
            try { sWeight = tmpTextBox.Name.Replace("txt", ""); } catch { sWeight = ""; }

            SqlConnection con = new SqlConnection(Global.ConnectionString);
            try
            {
                SqlCommand cmd = new SqlCommand("qryGetQuestionWeight", con);  // tblQuestion_Weighting
                cmd.Parameters.Clear();
                SqlParameter prm01 = cmd.Parameters.Add("@txtName", SqlDbType.NVarChar);
                prm01.Value = tmpTextBox.Name;
                SqlParameter prm02 = cmd.Parameters.Add("@Score", SqlDbType.Int);
                prm02.Value = Convert.ToInt16(tmpTextBox.Text);
                SqlParameter prm03 = cmd.Parameters.Add("@nPop", SqlDbType.NVarChar);
                prm03.Value = Global.Pop_Code_Name;
                cmd.CommandTimeout = Global.TimeOut;
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    WeightedScore = (dr["Weighted_Score"] is DBNull) ? 0 : Convert.ToInt32(dr["Weighted_Score"]);
                    tmpHDTextBox.Text = WeightedScore.ToString();
                }
                else if (dr.HasRows == false)
                {
                    MessageBox.Show("No corresponding weighting found for this previously selected field [" + sWeight + "]." + Environment.NewLine + "Please notify appropriate person immediately.", "Weighting", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                con.Close();
            }
            catch
            {
                con.Close();
            }
        }

        public bool CheckandAddRPD()
        {
            bool blnRtn = true;

            //blnRtn = CheckandSaveScoresData();

            blnRtn = CheckandSaveCustomerData();

            blnRtn = CheckandSaveAgentData();

            return blnRtn;
        }

        public bool CheckandSaveCustomerData()
        {
            bool blnCustomer = false;
            bool blnRtn = false;
                
            for (int j = 0; j < 27; j++)
            {
                switch (j)
                {
                    case 0:
                        if (cboPopCode.SelectedValue != GetDataContextValue("Pop").ToString()){ blnCustomer = true; }
                        break;
                    case 1:
                        if (cboSegment.Text != GetDataContextValue("Segment").ToString()) { blnCustomer = true; }
                        break;
                    case 2:
                        if (txtSurname.Text != GetDataContextValue("Surname").ToString()) { blnCustomer = true; }
                        break;
                    case 3:
                        if (txtFirstName.Text != GetDataContextValue("Firstname").ToString()) { blnCustomer = true; }
                        break;
                    case 4:
                        if (txtDOB.Text != GetDataContextValue("DOB").ToString()) { blnCustomer = true; }
                        break;
                    case 5:
                        if (txtDOD.Text != GetDataContextValue("DOD").ToString()) { blnCustomer = true; }
                        break;
                    case 6:
                        if (Convert.ToBoolean(GetDataContextValue("Deceased")) == false) // looking for "No" response
                        {
                            if(cboDeceased.Text.ToUpper() != "NO")
                            {
                                blnCustomer = true;
                            }
                        }
                        else  // looking for "Yes" response
                        {
                            if (cboDeceased.Text.ToUpper() != "YES")
                            {
                                blnCustomer = true;
                            }
                        }
                        break;
                    case 7:
                        if (txtDeselected.Text != GetDataContextValue("Deselected").ToString()) { blnCustomer = true; }
                        break;
                    case 8:
                        if (cboMarital.Text != GetDataContextValue("Marital").ToString()) { blnCustomer = true; }
                        break;
                    case 9:
                        if (cboGender.Text != GetDataContextValue("Gender").ToString()) { blnCustomer = true; }
                        break;
                    case 10:
                        if (txtPrivateAddress.Text != GetDataContextValue("MainAdd").ToString()) { blnCustomer = true; }
                        break;
                    case 11:
                        if (txtPostcode.Text != GetDataContextValue("MainPC").ToString()) { blnCustomer = true; }
                        break;
                    case 12:
                        if (txtSecondaryAddress.Text != GetDataContextValue("SecAdd").ToString()) { blnCustomer = true; }
                        break;
                    case 13:
                        if (cboResidence.Text != GetDataContextValue("Residence").ToString()) { blnCustomer = true; }
                        break;
                    case 14:
                        if (cboDomicile.Text != GetDataContextValue("Domicile").ToString()) { blnCustomer = true; }
                        break;
                    case 15:
                        if (cboOffice.Text != GetDataContextValue("Office").ToString()) { blnCustomer = true; }
                        break;
                    case 16:
                        if (cboTeam.Text != GetDataContextValue("Team").ToString()) { blnCustomer = true; }
                        break;
                    case 17:
                        if (cboCRMName.Text != GetDataContextValue("CRM_Name").ToString()) { blnCustomer = true; }
                        break;
                    case 18:
                        if (txtCRMDA.Text != GetDataContextValue("CRM_Appointed").ToString()) { blnCustomer = true; }
                        break;
                    case 19:
                        if (cboWealth.Text != GetDataContextValue("WealthLevel").ToString()) { blnCustomer = true; }
                        break;
                    case 20:
                        if (cboPathway.Text != GetDataContextValue("Pathway").ToString()) { blnCustomer = true; }
                        break;
                    case 21:
                        if (cboSource.Text != GetDataContextValue("Source").ToString()) { blnCustomer = true; }
                        break;
                    case 22:
                        if (cboSector.Text != GetDataContextValue("Sector").ToString()) { blnCustomer = true; }
                        break;
                    case 23:
                        if (cboLongTerm.Text != GetDataContextValue("LongTerm").ToString()) { blnCustomer = true; }
                        break;
                    case 24:
                        if (cboLifeEvents.Text != GetDataContextValue("LifeEvents").ToString()) { blnCustomer = true; }
                        break;
                    case 25:
                        if (txtNarrative.Text != GetDataContextValue("Narrative").ToString()) { blnCustomer = true; }
                        break;
                    case 26:
                        if (cboAllocatedTo.Text != GetDataContextValue("HNWUPID").ToString())
                        {
                            blnCustomer = true;
                            if(cboAllocatedTo.Text.Trim() == "")
                            {
                                blnRtn = true; // case can't be left unallocated
                            }
                            else
                            {
                                break;
                            }
                        } 
                        break;
                    default:
                        break;
                }
                if (blnCustomer == true) { break;}
            }

            if (blnCustomer == true) // save case as something has changed
            {
                // strQuery is qryAddCustomer or qryUpdateCustomer [qryUpdateCustomer now performs either] 
                try
                {
                    RPT.RPT_Data rpt = new RPT.RPT_Data(); // initialise data
                    rpt.SecAdd = txtSecondaryAddress.Text;
                    rpt.Strand = cboStrand.Text;
                    rpt.Segment = cboSegment.Text;
                    rpt.Surname = txtSurname.Text;
                    rpt.Firstname = txtFirstName.Text;
                    rpt.DOB = txtDOB.Text;
                    rpt.Deceased = (cboDeceased.Text.ToLower() == "yes") == true ? Convert.ToByte(1) : Convert.ToByte(0); /*convert Yes/No to byte*/
                    rpt.DOD = txtDOD.Text;
                    rpt.Deselected = txtDeselected.Text;
                    rpt.Marital = cboMarital.Text;
                    rpt.Gender = cboGender.Text;
                    rpt.MainAdd = txtPrivateAddress.Text;
                    rpt.MainPC = txtPostcode.Text;
                    rpt.SecAdd = txtSecondaryAddress.Text;
                    rpt.Residence = cboResidence.Text;
                    rpt.Domicile = cboDomicile.Text;
                    rpt.Office = cboOffice.Text;
                    rpt.Team = cboTeam.Text;
                    rpt.WealthLevel = cboWealth.Text;
                    rpt.Pathway = cboPathway.Text;
                    rpt.Source = cboSource.Text;
                    rpt.Sector = cboSector.Text;
                    rpt.LongTerm = cboLongTerm.Text;
                    rpt.LifeEvents = cboLifeEvents.Text;
                    rpt.Narrative = txtNarrative.Text;
                    rpt.HNWUPID = Convert.ToInt32(Global.PID);
                    rpt.UTR = Convert.ToDouble(txtUTR.Text);
                    rpt.Pop = cboPopCode.Text;
                    rpt.CRM_Name = cboCRMName.Text;
                    rpt.CRM_Appointed = txtCRMDA.Text;

                    rpt.UpdateCustomerData();
                    blnRtn = true;
                }
                catch
                {
                    blnRtn = false;
                }
            }
            else // don't save as nothing has changed
            {
                //leave as it as checks found no changes so no point in resaving data.
                blnRtn = true;
            }

            return blnRtn ;
        }

        public bool CheckandSaveAgentData()
        {
            bool blnAgent = false;
            bool blnRtn = false;
            string strTest = "";

            RPT.RPT_Data rpt = new RPT.RPT_Data(); // initialise data
            for (int j = 0; j < 8; j++)
            {
                switch (j)
                {
                    case 0:
                        strTest = GetDataContextValue("Agent") == null ? "" : GetDataContextValue("Agent").ToString();
                        if (txtFirm.Text != strTest) { blnAgent = true; }
                        break;
                    case 1:
                        if (cbo648.Text.ToUpper() != "NO") { blnAgent = true; }
                        break;
                    case 2:
                        strTest = GetDataContextValue("Surname") == null ? "" : GetDataContextValue("Surname").ToString();
                        if (txtSurname.Text != strTest) { blnAgent = true; }
                        break;
                    case 3:
                        strTest = GetDataContextValue("AgentCode") == null ? "" : GetDataContextValue("AgentCode").ToString();
                        if (txtAgentCode.Text != strTest) { blnAgent = true; }
                        break;
                    case 4:
                        strTest = GetDataContextValue("NamedAgent") == null ? "" : GetDataContextValue("NamedAgent").ToString();
                        if (txtNamedAgent.Text != strTest) { blnAgent = true; }
                        break;
                    case 5:
                        strTest = GetDataContextValue("OtherContact") == null ? "" : GetDataContextValue("OtherContact").ToString();
                        if (txtOtherContact.Text != strTest) { blnAgent = true; }
                        break;

                    case 6:
                        strTest = GetDataContextValue("AgentTelNo") == null ? "" : GetDataContextValue("AgentTelNo").ToString();
                        if (txtTelNo.Text != strTest) { blnAgent = true; }
                        break;
                    case 7:
                        strTest = GetDataContextValue("HNWUPID") == null ? "" : GetDataContextValue("HNWUPID").ToString();
                        if (cboChange.Text != strTest)
                        {
                            blnAgent = true;
                            if (cboAllocatedTo.Text.Trim() == "")
                            {
                                blnRtn = true; // case can't be left unallocated
                            }
                            else
                            {
                                break;
                            }
                        }
                        break;
                    default:
                        break;
                }
            }

            if (blnAgent == true) // save agent details as something has changed
            {
                // check no existing null record as crpd agent fields could already holds data if a record had been previosuly added but subsequently blanked out 
                string strA = GetDataContextValue("Agent") == null ? "" : GetDataContextValue("Agent").ToString();
                string strAC = GetDataContextValue("AgentCode") == null ? "" : GetDataContextValue("AgentCode").ToString();
                string strAA = GetDataContextValue("AgentAddress") == null ? "" : GetDataContextValue("AgentAddress").ToString();
                string strNA = GetDataContextValue("NamedAgent") == null ? "" : GetDataContextValue("NamedAgent").ToString();
                string strOC = GetDataContextValue("OtherContact") == null ? "" : GetDataContextValue("OtherContact").ToString();
                string strAT = GetDataContextValue("AgentTelNo") == null ? "" : GetDataContextValue("AgentTelNo").ToString();
                if ((strA == "") && (strAC == "") && (strAA == "") && (strNA == "") && (strOC == "") && (strAT == "") && (GetDataContextValue("Changed") == 0))
                {
                    rpt.ResetAgent();
                }
                try
                {
                    rpt.Agent = txtFirm.Text;
                    rpt.AgentCode = txtAgentCode.Text;
                    rpt.Agent648Held = (cbo648.Text.ToLower() == "yes") == true ? Convert.ToByte(1) : Convert.ToByte(0); /*convert Yes/No to byte*/
                    rpt.AgentAddress = txtAgentAddress.Text;
                    rpt.NamedAgent = txtNamedAgent.Text;
                    rpt.OtherContact = txtOtherContact.Text;
                    rpt.AgentTelNo = txtTelNo.Text;
                    rpt.Changed = (cboChange.Text.ToLower() == "yes") == true ? Convert.ToByte(1) : Convert.ToByte(0); /*convert Yes/No to byte*/
                    double dblUTR = Convert.ToDouble(txtUTR.Text.Trim());
                    rpt.UpdateAgentData(dblUTR);
                    blnRtn = true;
                }
                catch
                {
                    blnRtn = false;
                }
            }
            else // don't save as nothing has changed
            {
                //leave as it as checks found no changes so no point in resaving data.
                blnRtn = true;
            }

            return blnRtn;
        }

        public bool GetRPDDataFromForm()
        {
            bool blnRtn = false;
            RPT.RPT_Data rpt = new RPT.RPT_Data(); // initialise data
            rpt.Segment = cboSegment.Text;
            rpt.Surname = txtSurname.Text;
            rpt.Firstname = txtFirstName.Text;
            rpt.Gender = cboGender.Text;
            rpt.UTR = Convert.ToDouble(txtUTR.Text);
            rpt.Deceased = cboDeceased.Text.ToUpper() == "YES" ? Convert.ToByte(true) : Convert.ToByte(false);
            rpt.DOB = txtDOB.Text;
            rpt.DOD = txtDOD.Text;
            rpt.Deselected = txtDeselected.Text;
            rpt.Marital = cboMarital.Text;
            rpt.MainAdd  = txtPrivateAddress.Text;
            rpt.MainPC = txtPostcode.Text;
            rpt.SecAdd = txtSecondaryAddress.Text;
            rpt.Residence = cboResidence.Text;
            rpt.Domicile = cboDomicile.Text;
            rpt.Office = cboOffice.Text;
            rpt.Team = cboTeam.Text;
            rpt.HNWUPID = cboAllocatedTo.Text == "" ? 0 : Convert.ToInt32(System.Text.RegularExpressions.Regex.Match(cboAllocatedTo.Text, @"\d+").Value); 
            rpt.Agent = txtFirm.Text;
            rpt.Agent648Held = cbo648.Text.ToUpper() == "YES" ? Convert.ToByte(true) : Convert.ToByte(false);
            rpt.AgentCode = txtAgentCode.Text;
            rpt.AgentAddress = txtAgentAddress.Text;
            rpt.NamedAgent = txtNamedAgent.Text;
            rpt.OtherContact = txtOtherContact.Text;
            rpt.AgentTelNo = txtTelNo.Text;
            rpt.Changed = cboChange.Text.ToUpper() == "YES" ? Convert.ToByte(true) : Convert.ToByte(false);
            //rpt.OpenIDMS = Convert.ToInt16(txtOpenIDMS.Text);
            //rpt.ClosedIDMS = Convert.ToInt16(txtClosedIDMS.Text);
            rpt.Risks_Open = Convert.ToInt16(txtOpenRisks.Text);
            rpt.Settled_Risks = Convert.ToInt16(txtSettled.Text);
            rpt.Highest_Settlement = Convert.ToInt32(txtHighestSettled.Text);
            rpt.HPPenalty = Convert.ToInt16(txtHighestPercent.Text);
            switch (cboCurrentSuspensions.Text.ToUpper())
                {
                case "YES":
                    rpt.PSCurrent = 2;
                    break;
                case "NO":
                    rpt.PSCurrent = 1;
                    break;
                default:
                    rpt.PSCurrent = 0;
                    break;
                }
            switch (cboPrevSuspensions.Text.ToUpper())
            {
                case "YES":
                    rpt.PSPrevious = 2;
                    break;
                case "NO":
                    rpt.PSPrevious = 1;
                    break;
                default:
                    rpt.PSPrevious = 0;
                    break;
            }
            switch (cboFailures.Text.ToUpper())
            {
                case "YES":
                    rpt.PSFailures = 2;
                    break;
                case "NO":
                    rpt.PSFailures = 1;
                    break;
                default:
                    rpt.PSFailures = 0;
                    break;
            }
            rpt.WealthLevel = cboWealth.Text;
            rpt.Pathway = cboPathway.Text;
            rpt.Source = cboSource.Text;
            rpt.Sector = cboSector.Text;
            rpt.LongTerm = cboLongTerm.Text;
            rpt.LifeEvents = cboLifeEvents.Text;
            rpt.Narrative = txtNarrative.Text;
            try { rpt.QSScore = Convert.ToInt16(txtQSScore.Text); } catch { rpt.QSScore =  0; };
            try { rpt.RPTPRScore = Convert.ToInt16(txtRSScore.Text); } catch { rpt.RPTPRScore = 0; };
            try { rpt.RPTAVScore = Convert.ToInt16(txtAVScore.Text); } catch { rpt.RPTAVScore = 0; };
            try { rpt.CGScore = Convert.ToInt16(txtCGScore.Text); } catch { rpt.CGScore = 0; };
            try { rpt.ResScore = Convert.ToInt16(txtRESScore.Text); } catch { rpt.ResScore = 0; };
            try { rpt.CRMScore = Convert.ToInt16(txtCRMScore.Text); } catch { rpt.CRMScore = 0; };
            try { rpt.PriorityScore = Convert.ToInt16(txtPRScore.Text); } catch { rpt.PriorityScore = 0; };
            try { rpt.Percentile = Convert.ToInt32(txtPercentile.Text); } catch { rpt.Percentile = 0; };
            rpt.Pop = cboPopCode.Text;
            rpt.CRMExplanation = txtCRMExplanation.Text;
            return blnRtn;
        }

        public dynamic GetDataContextValue(string sPropertyName)
        {
            dynamic vRtn = 0;
            try
            {
                vRtn = this.DataContext.GetType().GetProperty(sPropertyName).GetValue(this.DataContext, null);
            }
            catch
            {
                vRtn = 0;
            }
            return vRtn;
        }

        private void CboTeam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((bFormLoaded == true))
            {
                bool blnTest = PopulateAllocatedToCombo();
                cboAllocatedTo.Text = "";
                return;
            }
            else
            {
            }
                //bool blnKeepCRMDA = false;
                //if (cboTeam.Text == "")
                //{
                //    //cboAllocatedTo.Claer;
                //    cboAllocatedTo.Text = "";
                //    return;
                //}

                //if(txtTeam.Text == "")
                //{
                //    if(txtTeam.Text == cboTeam.Text)
                //    {
                //        return;
                //    }
                //}

                //PopulateAllocatedToCombo();
                //PopulateCRMCombo();

                //blnKeepCRMDA = false;

                //bool itemExists = false;
                //foreach (ComboBoxItem cbi in cboCRMName.Items)
                //{
                //    itemExists = cbi.Content.Equals(txtCRMName.Text);
                //    if (itemExists)
                //    {
                //        blnKeepCRMDA = true;
                //        break;
                //    }
                //}

                //if(blnKeepCRMDA == false)
                //{
                //    txtCRMDA.Text = "";
                //}

                //txtTeam.Text = cboTeam.Text;
                //cboAllocatedTo.IsEnabled = true;

            }

        private void CmdUpdateClose_Click(object sender, RoutedEventArgs e)
        {

            cmdUpdateClose.IsEnabled = false;
            Globals.blnUpdate = false;
            RecalculateBehaviours();
            SaveRecord();

            if(Globals.blnUpdate == true)
            {
                CmdCancel_Click(sender, e);
            }
            else
            {
                cmdUpdateClose.IsEnabled = true;
            }

        }

        private void RecalculateResults()
        {
            int PSScore = 0;


            if (IsParseable(txtQSScore.Text.ToString(), false) == false)
            {
                return;
            }
            if (IsParseable(txtAVScore.Text.ToString(), false) == false)
            {
                return;
            }
            if (IsParseable(txtRSScore.Text.ToString(), false) == false)
            {
                return;
            }
            if (IsParseable(txtCGScore.Text.ToString(), false) == false)
            {
                return;
            }
            if (IsParseable(txtRESScore.Text.ToString(), false) == false)
            {
                return;
            }
            if (IsParseable(txtCRMScore.Text.ToString(), true) == false)
            {
                return;
            }

            PSScore = int.Parse(txtQSScore.Text.ToString()) + int.Parse(txtAVScore.Text.ToString()) + int.Parse(txtRSScore.Text.ToString()) + int.Parse(txtCGScore.Text.ToString()) + int.Parse(txtRESScore.Text.ToString()) + int.Parse(txtCRMScore.Text.ToString());

            txtPRScore.Text = Convert.ToString(PSScore);
        }

        private bool IsParseable(string strText, bool blnCRM)
        {
            int number;

            bool ParseableCheck = Int32.TryParse(strText, out number);

            if (ParseableCheck)
            {
                return true;
            }
            else 
            {
                if (blnCRM == true)
                {
                    txtCRMScore.Text = "0";
                }
                else
                { 
                    MessageBox.Show("Unable to calculate Priority Score.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                return false;
            }
        }

        private void TxtQSScore_LostFocus(object sender, RoutedEventArgs e)
        {
            RecalculateResults();

            // replot graph
        }

        private void PgResults_GotFocus(object sender, RoutedEventArgs e)
        {
            RecalculateBehaviours();
        }

        //private void CmdUpdateClose_Click(object sender, RoutedEventArgs e)
        //{
        //    RecalculateBehaviours();

        //    BindingExpression obj = txtSecondaryAddress.GetBindingExpression(TextBox.TextProperty);
        //    obj.UpdateSource();
        //}

        private void ReplotChart()
        {
            
        }

        private void RecalculateBehaviours()
        {
            int iCSuspensions;
            int iPSuspensions;
            int iFailures;
            int iCRMBoost;
            int iYear = 2000;
            var vSegment = "";
            int iChartPoints;

            if (this.cboCurrentSuspensions.SelectedItem == null)
            {
                iCSuspensions = 0;
            }
            else
            {
                iCSuspensions = Convert.ToInt16(this.cboCurrentSuspensions.SelectedItem.ToString());
            }

            if (this.cboPrevSuspensions.SelectedItem == null)
            {
                iPSuspensions = 0;
            }
            else
            {
                iPSuspensions = Convert.ToInt16(this.cboPrevSuspensions.SelectedItem.ToString());
            }

            if (this.cboFailures.SelectedItem == null)
            {
                iFailures = 0;
            }
            else
            {
                iFailures = Convert.ToInt16(this.cboFailures.SelectedItem.ToString());
            }

            if (this.txtCRMScore.Text == "")
            {
                iCRMBoost = 0;
            }
            else
            {
                iCRMBoost = Convert.ToInt16(this.txtCRMScore.Text.ToString());
            }

            try { iYear = Convert.ToInt16(this.lblPopYear.Text); } catch { }

            DataTable dt = new DataTable("dgRank");

            try
            {
                // Connecting the SQL Server
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                con.Open();
                // Calling the Stored Procedure
                SqlCommand cmd = new SqlCommand("qryGetBehaviourScoresandRank", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@nOpenRisks", SqlDbType.Int).Value = Convert.ToInt16(this.txtOpenRisks.Text.ToString());
                cmd.Parameters.Add("@nClosedRisks", SqlDbType.Int).Value = Convert.ToInt16(this.txtSettled.Text.ToString());
                cmd.Parameters.Add("@nHighestSettlement", SqlDbType.Float).Value = Convert.ToDouble(this.txtHighestSettled.Text.ToString());
                cmd.Parameters.Add("@nHPPenalty", SqlDbType.Float).Value = Convert.ToDouble(this.txtHighestPercent.Text.ToString());
                cmd.Parameters.Add("@nCurrentSuspensions", SqlDbType.Int).Value = iCSuspensions;
                cmd.Parameters.Add("@nPreviousSuspensions", SqlDbType.Int).Value = iPSuspensions;
                cmd.Parameters.Add("@nSuspensionFailures", SqlDbType.Int).Value = iFailures;
                cmd.Parameters.Add("@nOpenIDMS", SqlDbType.Int).Value = Convert.ToInt16(this.txtOpenIDMS.Text.ToString());
                cmd.Parameters.Add("@nClosedIDMS", SqlDbType.Int).Value = Convert.ToInt16(this.txtClosedIDMS.Text.ToString());
                cmd.Parameters.Add("@nAvoidanceScore", SqlDbType.Int).Value = Convert.ToInt16(this.txtAVScore.Text.ToString());
                cmd.Parameters.Add("@nRiskScore", SqlDbType.Int).Value = Convert.ToInt16(this.txtRSScore.Text.ToString());
                cmd.Parameters.Add("@nCGScore", SqlDbType.Int).Value = Convert.ToInt16(this.txtCGScore.Text.ToString());
                cmd.Parameters.Add("@nResScore", SqlDbType.Int).Value = Convert.ToInt16(this.txtRESScore.Text.ToString());
                cmd.Parameters.Add("@nCRMScore", SqlDbType.Int).Value = iCRMBoost;
                cmd.Parameters.Add("@nCurrentPSScore", SqlDbType.Int).Value = Convert.ToInt16(this.txtPRScore.Text.ToString());
                cmd.Parameters.Add("@nCurrentRank", SqlDbType.Float).Value = Convert.ToDouble(this.txtPercentile.Text.ToString());
                cmd.Parameters.Add("@nCalendarYear", SqlDbType.Int).Value = iYear;
                cmd.Parameters.Add("@nUTR", SqlDbType.Int).Value = Convert.ToInt32(this.txtUTR.Text.ToString());
                cmd.Parameters.Add("@nPop", SqlDbType.Text).Value = this.cboPopCode.SelectedValue.ToString();


                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);

                this.txtQSScore.Text = dt.Rows[0]["NewQSScore"].ToString();
                this.txtPRScore.Text = dt.Rows[0]["NewPSScore"].ToString();
                this.txtPercentile.Text = dt.Rows[0]["NewRank"].ToString();
                vSegment = dt.Rows[0]["NewSegment"].ToString();
                if (vSegment == "")
                {
                    vSegment = "High Risk";
                }
                this.cboSegment.SelectedValue = vSegment.Trim();

                try  //Update Graph
                {
                    iChartPoints = Globals.dtGraph.Rows.Count - 1;
                    Globals.dtGraph.Rows[iChartPoints]["Ranking"] = dt.Rows[0]["NewRank"].ToString();
                    this.mscHistory.ItemsSource = null;
                    this.mscHistory.ItemsSource = Globals.dtGraph.DefaultView;

                    //Update Grid
                    Globals.dtGrid.Rows[0]["Priority Score"] = dt.Rows[0]["NewPSScore"].ToString();
                    Globals.dtGrid.Rows[0]["Ranking"] = dt.Rows[0]["NewRank"].ToString();
                    Globals.dtGrid.Rows[0]["Segment"] = dt.Rows[0]["NewSegment"].ToString();
                    this.dgHistorical.ItemsSource = null;
                    this.dgHistorical.ItemsSource = Globals.dtGrid.DefaultView;
                }
                catch
                {

                }
            }
            catch
            {
                System.Windows.MessageBox.Show("Unable to recalculate score.  Press ReCheck Ranking Button to try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cmdPercentile_Click(object sender, RoutedEventArgs e)
        {
            RecalculateBehaviours();
        }
    }

}
