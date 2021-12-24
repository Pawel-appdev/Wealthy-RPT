using Microsoft.Win32;
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

using System.IO;
using System.Data.OleDb;


namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for ImportData.xaml
    /// </summary>
    public partial class ImportData : Window
    {
        public static string strTable;
        public static string strTableTop;

        public ImportData()
        {
            InitializeComponent();

            PopulateData();
        }

        private void PopulateData()
        {
            try
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                SqlCommand cmd = new SqlCommand("qryGetAdditionalDataSourceInfo", con);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds, "SourceInfo");
                cboData.ItemsSource = ds.Tables[0].DefaultView;
                cboData.DisplayMemberPath = ds.Tables[0].Columns["Friendly_Name"].ToString();
                cboData.SelectedValuePath = ds.Tables[0].Columns["Default_Location"].ToString();

                dgData.ItemsSource = ds.Tables[0].DefaultView;
            }
            catch
            {
                MessageBox.Show("Unable to connect to database.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void cboData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            txtSource.Text = cboData.SelectedValue.ToString();
        }

        private void btnDelPrevData_Click(object sender, RoutedEventArgs e)

        {
            //string strTableTop;

            if (txtSource.Text == "")
            {
                MessageBox.Show("Please confirm the data to be deleted.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (cboData.Text == "")
            {
                MessageBox.Show("There doesn't appear to be a Table Instruction for this Data." +"\n" + "Please report to the Support Team Mailbox. See Help/About.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            else
            {

                GetTableName();

                int intRight = strTable.IndexOf(" ");

                //MessageBox.Show(strTable.ToString());
                
                if(intRight != -1)
                {
                    strTable = strTable.Substring(0, intRight);
                    
                }
                var selectedOption = MessageBox.Show("Are you sure you wish to delete the existing Data.", "Weathy RPT", MessageBoxButton.YesNo, MessageBoxImage.Question);
                switch (selectedOption)
                {
                    case MessageBoxResult.Yes:
                        DeleteYear  del_Year = new DeleteYear();
                        del_Year.ShowDialog();
                        break;
                    case MessageBoxResult.No:
                        break;

                }
            }
        }

        private void btnSource_Click(object sender, RoutedEventArgs e)
        {
            if (txtSource.Text == "")
            {
                MessageBox.Show("Please confirm Data Source first.","Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            //extract folder
            string strDirectory = System.IO.Path.GetDirectoryName(txtSource.Text);

            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Please select the Source File from which to import the Data.",
                //InitialDirectory = strDirectory,
                FileName = "", 
                DefaultExt = ".xlsx",
                Filter = "Excel files (.xlsx)|*.xlsx" 
            };

            // Show open file dialog box
            Nullable<bool> result = openFileDlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Select document
                string filename = openFileDlg.FileName;
                if(filename != "")
                {
                    MessageBox.Show("If you would like this location to be the default location in future," + "\n" + "Please update the details using the Update DB Tables Function.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }

        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            string strImportSQL;
            string strImportLocation;
            string strIdentifier;

            if (txtSource.Text == "")
            {
                MessageBox.Show("Please confirm the data to be imported.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (cboData.Text == "")
            {
                MessageBox.Show("There doesn't appear to be a Table Instruction for this Data." + "\n" + "Please report to the Support Team Mailbox. See Help/About.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int intDataIndex = cboData.SelectedIndex;
            //get Additional_Data_Import_Instruction corresponding to Friendly Name
            strImportSQL = (dgData.Items[intDataIndex] as DataRowView).Row.ItemArray[5].ToString();

            if (txtSource.Text == "")
            {
                MessageBox.Show("Please confirm the source location.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            else
            {
                strImportLocation = txtSource.Text;
                //delete after testing on Dev machine
                //string fileName = System.IO.Path.GetFileName(strImportLocation);
                //strImportLocation = @"C:\Users\AT031564\Import_Test\" + fileName;
            }

            //get AutoNumber_Identifier corresponding to Friendly Name
            strIdentifier  = (dgData.Items[intDataIndex] as DataRowView).Row.ItemArray[7].ToString();

            if (strIdentifier == "")
            {
                MessageBox.Show("There doesn't appear to be a specifiec Auto Number Data Indentifier for the Data Table." + "\n" + "Please report to the Support Team Mailbox. See Help/About.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }


            GetTableName();
            //MessageBox.Show(strTable.ToString());


            if (File.Exists (strImportLocation))
            {
                try
                {
                    if (GetRowsCount(strTable) > 0)
                    {
                        MessageBox.Show("The Table Top 0 Line Instruction has returned some records." + "\n" + "Please report to the Support Team Mailbox. See Help/About.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    //Import from Excel

                    string TableName = strTable;
                    String SchemaName = @"dbo";
                    string fileFullPath = strImportLocation;

                    //Create Connection to SQL Server Database 
                    SqlConnection SQLConnection = new SqlConnection(Global.ConnectionString);

                    //Create Excel Connection
                    string ConStr;
                    string HDR;
                    HDR = "YES";
                    //ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    //    + fileFullPath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
                    ConStr = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source="
                        + fileFullPath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
                    OleDbConnection cnn = new OleDbConnection(ConStr);


                    //Get data from Excel Sheet to DataTable
                    OleDbConnection Conn = new OleDbConnection(ConStr);
                    Conn.Open();
                    OleDbCommand oconn = new OleDbCommand(strImportSQL, Conn);
                    OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    Conn.Close();

                    SQLConnection.Open();
                    //Load Data from DataTable to SQL Server Table.
                    using (SqlBulkCopy BC = new SqlBulkCopy(SQLConnection))
                    {
                        BC.DestinationTableName = SchemaName + "." + TableName;
                        foreach (var column in dt.Columns)
                            BC.ColumnMappings.Add(column.ToString(), column.ToString());
                        BC.WriteToServer(dt);
                    }
                    SQLConnection.Close();
                    MessageBox.Show("Data import complete.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                MessageBox.Show("Check the Source Location as file doesn't appear to exist.", "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }





        }
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void GetTableName()
        {
            //string strTableTop;

            int intDataIndex = cboData.SelectedIndex;
            //get strTableTopLineScript corresponding to Friendly Name
            strTableTop = (dgData.Items[intDataIndex] as DataRowView).Row.ItemArray[6].ToString();

            // Remainder of string starting at "FROM"
            int intLeft = strTableTop.IndexOf("FROM");
            strTable = strTableTop.Substring(intLeft + 4);
            strTable = strTable.Trim();
            int intRight = strTable.IndexOf(" ");

            if (intRight != -1)
            {
                strTable = strTable.Substring(0, intRight);
            }
        }

        public static int GetRowsCount(string tablename)
        {
            //string stmt = string.Format("SELECT COUNT(*) FROM {0}", tablename);
            
            string connStr = Global.ConnectionString;
            int count = 0;

            //MessageBox.Show(strTableTop, "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Question);
            try
            {
                using (SqlConnection thisConnection = new SqlConnection(connStr))
                {
                    thisConnection.Open();
                    var com = thisConnection.CreateCommand();
                    com.CommandText = strTableTop + ";Select @@RowCount;";
                    using (var reader = com.ExecuteReader())
                    {
                        reader.NextResult();
                        if(reader.Read())
                        {
                            count = (int)reader[0];
                            //MessageBox.Show(count.ToString(), "Customer Management Tool", MessageBoxButton.OK, MessageBoxImage.Question);
                        }
                    }
                }
                //{
                //    using (SqlCommand cmdCount = new SqlCommand(strTableTop, thisConnection))
                //    {
                //        thisConnection.Open();
                //        count = (int)cmdCount.ExecuteScalar();
                //    }
                //}
                return count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return 1;  // amended to 1 so instruction doesn't complete as something has gone wrong above. 25/11/2020
            }
        }

    }
}
