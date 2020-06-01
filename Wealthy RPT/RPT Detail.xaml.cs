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
            if(Globals.blnAccess == true)
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

        private void TxtSettledRisks_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(txtSettledRisks);
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

        private void LvwPrevRes_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Globals.blnAccess == true)
            {
                ShowActiveControl(lvwPrevRes);
            }
        }

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
                        if(item.ToString() == txtOffice.Text)
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
            //RecalculateResults();
        }

        private void CmdSave_Click(object sender, RoutedEventArgs e)
        {
            BindingExpression obj = txtSecondaryAddress.GetBindingExpression(TextBox.TextProperty);
            obj.UpdateSource();
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
            if(cboOffice.Text == "")
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
    }

}
