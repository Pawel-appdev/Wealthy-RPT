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
//using System.Windows.Forms;
//using MessageBox = System.Windows.MessageBox;

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
            InitializeComponent();
            PopulateCombos();

            this.WindowState = WindowState.Maximized;

        }


        private void PopulateCombos()
        {
            try
            {
                Lookups lu = new Lookups();
                lu.GetRPTDetailLookups();
                lu.GetRPTDetailOfficeCRMs();
                //lu.GetOfficeCRMs();

                // == Customer Segment

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
                cboDeceased.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboDeceased.DisplayMemberPath = "Options";
                cboDeceased.SelectedValuePath = "DecodedValue";

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

                // cboOffice populated on Window_Loaded() in order to pass selected PopCode

                // cboTeam  populated on Window_Loaded() in order to pass selected Office and PopCode

                // cboAllocatedTo: qryGetOfficeTeamStaff @nOffice @nTeam GetOfficeTeamStaff(string strOffice, string strTeam)
                //cboAllocatedTo.ItemsSource = lu.dsOfficeTeams.Tables[0].DefaultView;
                //cboAllocatedTo.DisplayMemberPath = "Team Identifier";
                //cboAllocatedTo.SelectedValuePath = "Team Identifier";

                //cboCRMName
                cboCRMName.ItemsSource = lu.dsRPTDetailOfficeCRMs.Tables[0].DefaultView;
                cboCRMName.DisplayMemberPath = "CRM_Name";
                cboCRMName.SelectedValuePath = "CRM_Name";

                // == Appointed Agent
                //cbo648 [yes/no]
                cbo648.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cbo648.DisplayMemberPath = "Options";
                cbo648.SelectedValuePath = "DecodedValue";
                //cboChange [yes/no]
                cboChange.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboChange.DisplayMemberPath = "Options";
                cboChange.SelectedValuePath = "DecodedValue";

                // == Behaviors
                //cboCurrentSuspensions [yes/no]
                cboCurrentSuspensions.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboCurrentSuspensions.DisplayMemberPath = "Options";
                cboCurrentSuspensions.SelectedValuePath = "DecodedValue";
                //cboPrevSuspensions [yes/no]
                cboPrevSuspensions.ItemsSource = lu.dsRPTDetailCombo.Tables[12].DefaultView;
                cboPrevSuspensions.DisplayMemberPath = "Options";
                cboPrevSuspensions.SelectedValuePath = "DecodedValue";
                //cboFailures [yes/no]
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

            //' questionnaire scores     [behaviours]
            //Me.txtOpenIDMS.Text = 0
            //Me.txtClosedIDMS.Text = 0
            //Me.txtOpenRisks.Text = 0
            //Me.txtSettled.Text = 0
            //Me.txtHighestSettled.Text = 0
            //Me.txtHighestPercent.Text = 0

            //## frmRPD  chkCRMDescretion_Click()

            //## modGetData  GetCRM

            //## modGlobal
            //## Public gs_CRM(2) As Integer

            // If gs_CRM(1) = 0 Or gs_CRM(2) = 0 Then
            //  GetCRM
            // End If

            //TabFrames
            //this.Activated += AfterLoading;
            bFormLoaded = true;
        }

        //private void AfterLoading(object sender, EventArgs e)
        //{
        //    this.Activated -= AfterLoading;
        //    bFormLoaded = true;
        //}

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
                    cboOffice.Items.Clear();
                    bool blnTest = PopulateOfficeCombo();
                    foreach (var item in cboOffice.Items)
                    {
                        if(item.ToString() == txtOffice.Text)
                        {
                            cboOffice.Text = txtOffice.Text;
                            break;
                        }
                    }
                    cboTeam.Items.Clear();
                }
            }
        }
    }

}
