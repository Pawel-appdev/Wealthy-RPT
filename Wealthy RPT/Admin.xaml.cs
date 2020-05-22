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

        //private static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        //private static string SQLquery = "SELECT * FROM [tblAdditional_Data_Sources]";

        //private static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        //private static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        //private static DataTable dt = new DataTable("Tables");

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
            cmdBuilder.dt.Clear();
            cmdBuilder.sda.Fill(cmdBuilder.dt);
            dgTables.ItemsSource = cmdBuilder.dt.DefaultView;


            ////try
            ////{
            ////SqlConnection conn = new SqlConnection(Global.ConnectionString);
            ////SqlDataAdapter da = new SqlDataAdapter(strAddlDataInstr, conn);
            ////DataSet ds = new DataSet();
            ////da.Fill(ds, "Tables");

            ////dgTables.ItemsSource = ds.Tables[0].DefaultView;
            ////}
            ////catch (Exception ex)
            ////{
            ////    MessageBox.Show(ex.ToString());
            ////}

            //string strSQLquery = strAddlDataInstr;// "SELECT * FROM [tblAdditional_Data_Sources]";
            //DataTable dt = new DataTable();
            //SqlConnection connection = new SqlConnection(Global.ConnectionString);
            //connection.Open();
            //SqlDataAdapter sqlDa = new SqlDataAdapter();
            //sqlDa.SelectCommand = new SqlCommand(strSQLquery, connection);
            //SqlCommandBuilder cb = new SqlCommandBuilder(sqlDa);
            //sqlDa.Fill(dt);
            //dt.Rows[0]["Friendly_Name"] = "Info 1";
            //dt.Rows[1]["Calendar_Year"] = "";
            //sqlDa.UpdateCommand = cb.GetUpdateCommand();
            //sqlDa.Update(dt);
            //dgTables.ItemsSource = dt.DefaultView;
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            cmdBuilder.conn.Open();
            SqlCommandBuilder builder = new SqlCommandBuilder(cmdBuilder.sda);
            cmdBuilder.sda.UpdateCommand = builder.GetUpdateCommand();
            cmdBuilder.sda.Update(cmdBuilder.dt);
            cmdBuilder.conn.Close();
        }

        private void dgTables_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            cmdBuilder.conn.Open();
            SqlCommandBuilder builder = new SqlCommandBuilder(cmdBuilder.sda);
            cmdBuilder.sda.UpdateCommand = builder.GetUpdateCommand();
            cmdBuilder.sda.Update(cmdBuilder.dt);
            cmdBuilder.conn.Close();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
