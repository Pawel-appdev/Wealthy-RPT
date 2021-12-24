using System;
using System.Collections.Generic;
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
    /// Interaction logic for DeleteYear.xaml
    /// </summary>
    public partial class DeleteYear : Window
    {
        public string strDelYear;

        public DeleteYear()
        {
            InitializeComponent();

            PopulateDeleteYearCombo();
        }

        private void PopulateDeleteYearCombo()
        {
            cboDeleteYear.Items.Add("All");
            //cboYear
            int intCurrentYear = DateTime.Today.Year;
            for (int intYear = intCurrentYear - 6; intYear <= intCurrentYear; intYear++)
            {
                cboDeleteYear.Items.Add(intYear);
            }
        }

        private void cboDeleteYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            strDelYear = cboDeleteYear.SelectedValue.ToString();
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string strDelete;
                int intYearToDelete;

                if (int.TryParse(strDelYear, out intYearToDelete))//delete particular year unless this is tblInfo
                {

                    if(ImportData.strTable.IndexOf("tblInfo") != -1)
                    {
                        MessageBox.Show("Unable to delete any particular year." + "\n" + "Select 'All' if you wish to delete all records.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    if (ImportData.strTable.IndexOf("Enquiry_Data") != -1)
                    {
                        strDelete = "DELETE FROM " + ImportData.strTable + " WHERE Enq_year = " + intYearToDelete;
                    }
                    else
                    {
                        strDelete = "DELETE FROM " + ImportData.strTable + " WHERE CalendarYear = " + intYearToDelete;
                    }
                }
                else//delete all years
                {
                    if (string.IsNullOrEmpty(strDelYear))
                    {
                        MessageBox.Show("A year has not been selected.", "Missing parameter", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    strDelete = "DELETE FROM " + ImportData.strTable;
                }

                SqlConnection con = new SqlConnection(Global.ConnectionString);
                con.Open();

                string sql = strDelete;
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
                
                this.Close();
                MessageBox.Show("Data removal process completed.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (SystemException ex)
            {
                MessageBox.Show(string.Format( ex.Message),"Error");
            }

        }
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
