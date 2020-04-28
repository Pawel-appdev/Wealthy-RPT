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


            //get Additional_Data_Instruction corresponding to Friendly Name
            strAddlDataInstr = (dgQueries.Items[intTableIndex] as DataRowView).Row.ItemArray[1].ToString();

            //get Main_Key corresponding to Friendly Name
            strMainKey = (dgQueries.Items[intTableIndex] as DataRowView).Row.ItemArray[2].ToString();

            FillTablesGrid(strAddlDataInstr);
        }

        private void FillTablesGrid(string strAddlDataInstr)
        {
            try
            {
            SqlConnection conn = new SqlConnection(Global.ConnectionString);
            SqlDataAdapter da = new SqlDataAdapter(strAddlDataInstr, conn);
            DataSet ds = new DataSet();
            da.Fill(ds, "Tables");

            dgTables.ItemsSource = ds.Tables[0].DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
