using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
    /// Interaction logic for Email.xaml
    /// </summary>
    public partial class Email : Window

    {
        public double dUTR;
        public string Email_ID;
        public double dPercentile;
        public string sPop;

        public Email(double dblUTR,string EmailID, string strContact, string strEmail, string strRole, string strDateAdded, double dblPercentile, String strPop)
        {
            InitializeComponent();
            //testAction();
            dUTR = dblUTR;
            Email_ID = EmailID;
            dPercentile = dblPercentile;
            sPop = strPop;

            this.txtContact.Text = strContact;
            this.txtEmailAddr.Text = strEmail;
            this.txtRole.Text = strRole;

            if(strDateAdded == "")//button 'Add' clicked
            {
                txtDateAdded.Text = "";
            }
            else//button 'Update' clicked
            {
                this.lblDate.Content = Convert.ToDateTime(strDateAdded);
            }
        }

        private void addEmail()
        {

            if(txtContact.Text.Trim() == "" || txtEmailAddr.Text.Trim() == "" || txtRole.Text.Trim() == "")
            {
                MessageBox.Show("None of the required fields appear to have been completed." + "\n" + "\n" + "Please rectify.", "Wealthy RPT", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            try
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                con.Open();

                SqlCommand cmd = new SqlCommand("qryAddEmailAddress", con);
                cmd.CommandType = CommandType.StoredProcedure;
                
                cmd.Parameters.Add("@nContact", SqlDbType.Text).Value = this.txtContact.Text;
                cmd.Parameters.Add("@nEmailAddress", SqlDbType.Text).Value = this.txtEmailAddr.Text;
                cmd.Parameters.Add("@nContactRole", SqlDbType.Text).Value = this.txtRole.Text;
                cmd.Parameters.Add("@nDateAdded", SqlDbType.DateTime).Value = DateTime.Now.ToString("dd/MM/yyyy");
                cmd.Parameters.Add("@nUTR", SqlDbType.Float).Value = dUTR;
                
                cmd.ExecuteNonQuery();
                con.Close();

                this.Close();
                int intTab = 0;
                Forms.reloadForm(dUTR, intTab, dPercentile, sPop);

        }
            catch
            {
                MessageBox.Show("Contact details have not been added", "Wealthy RPT", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        private void updateEmail()
        {
            try
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                con.Open();
                SqlCommand cmd = new SqlCommand("qryUpdateEmailAddress", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@oEmailID", SqlDbType.Int).Value = Email_ID;
                cmd.Parameters.Add("@nContact", SqlDbType.Text).Value = this.txtContact.Text;
                cmd.Parameters.Add("@nEmailAddress", SqlDbType.Text).Value = this.txtEmailAddr.Text;
                cmd.Parameters.Add("@nRole", SqlDbType.Text).Value = this.txtRole.Text;
                cmd.ExecuteNonQuery();
                con.Close();

                this.Close();
                int intTab = 0;
                Forms.reloadForm(dUTR, intTab, dPercentile, sPop);
        }
            catch
            {
                MessageBox.Show("Contact details have not been updated", "Wealthy RPT", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        private void btnAction_Click(object sender, RoutedEventArgs e)
        {
            if (btnAction.Content.ToString() == "Add")
            {
                addEmail();
            }

            if (btnAction.Content.ToString() == "Update")
            {
                updateEmail();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
