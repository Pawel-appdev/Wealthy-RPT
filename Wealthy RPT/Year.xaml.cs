using System;
using System.Windows;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for Year.xaml
    /// </summary>
    public partial class Year : Window
    {

        public static string strReportYear;
        private static string strYear;

        public Year()
        {
            InitializeComponent();
            PopulateCombo();
        }

        private void PopulateCombo()
        {
            int intCurrentYear = DateTime.Today.Year;
            for (int intYear = intCurrentYear - 5; intYear <= intCurrentYear; intYear++)
            {
                cboReportYear.Items.Add(intYear);
            }
            //cboReportYear.SelectedValue = intCurrentYear;
        }

        private void cboReportYear_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                strYear = cboReportYear.SelectedValue.ToString();
            }
            catch
            {
                MessageBox.Show("You have not selected any year","No year selected",MessageBoxButton.OK ,MessageBoxImage.Error);
                return;
            }
            
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            strReportYear = "yyyy";
            this.Close();
        }

        private void btnSubmit_Click(object sender, RoutedEventArgs e)
        {
            strReportYear = strYear;
            if(string.IsNullOrEmpty(strReportYear))//no year selected
            {
                MessageBox.Show("Select a year or click 'Cancel' button to close this form", "No year selected", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else 
            {
                this.Close();
            }

        }
    }
}
