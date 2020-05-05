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

        public Email(double dblUTR)
        {
            InitializeComponent();
            //testAction();
            dUTR = dblUTR;
        }

        //private void testAction()
        //{
        //    switch (this.Title)
        //    {
        //        case "Add Email Address":
        //            break;

        //        case "Update Email Address":
        //            break;

        //        case "Delete Email Address":
        //            break;
        //    }

        //}

        private bool addEmail()
        {
            //try
            //{
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

                refreshScreen();

                return true;
            //}
            //catch
            //{
            //    return false;
            //}
        }

        private void btnAction_Click(object sender, RoutedEventArgs e)
        {
            if (btnAction.Content.ToString() == "Add")
            {
                addEmail();
            }
        }

        private void refreshScreen()
        {
            RPT_Detail.refreshScreen(dUTR);
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
