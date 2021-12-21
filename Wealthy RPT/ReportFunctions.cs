using CsvHelper;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExPat_UI.Reports
{
    class ReportFunctions
    {

        /// <summary> DataTableToCSV
        /// Converts DataTable to a CSV output 
        /// </summary>
        /// <param name="filePath">Path to Export file to.</param>
        /// <param name="fileName">Filename to Export file to.</param>
        /// <param name="dataTable">DataTable to be exported.</param>
        /// <returns></returns>
        public void DataTableToCSV(string filePath, string fileName, DataTable dataTable)
        {
            filePath += @"" + fileName + " " + DateTime.Now.ToString("yyyy-MM-dd") + ".csv";
            using (StreamWriter writer = new StreamWriter(filePath))
            using (CsvWriter csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                // Write columns
                foreach (DataColumn column in dataTable.Columns)
                {
                    csv.WriteField(column.ColumnName);
                }
                csv.NextRecord();

                // Write row values
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        csv.WriteField(row[i]);
                    }
                    csv.NextRecord();
                }
            }
        }

        /// <summary> DataGridtoDataTable
        /// Takes passed DataGrid, splits it up and writes out Columns
        /// and then rows
        /// </summary>
        /// <param name="dataGrid">DataGrid</param>
        /// <returns>DataTable</returns>
        public DataTable DataGridtoDataTable(DataGrid dataGrid)
        {
            dataGrid.SelectAllCells();
            dataGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, dataGrid);
            dataGrid.UnselectAllCells();
            String result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
            string[] Lines = result.Split(new string[] { "\r\n", "\n\n" }, StringSplitOptions.None);
            string[] Fields;
            Fields = Lines[0].Split(new char[] { ',' });
            int Cols = Fields.GetLength(0);

            DataTable dataTable = new DataTable();
            for (int i = 0; i < Cols; i++)
            {
                dataTable.Columns.Add(Fields[i].ToUpper(), typeof(string));
            }
            DataRow Row;
            var xx = Lines.GetLength(0) - 1;

            for (int i = 1; i < Lines.GetLength(0) - 1; i++)
            {
                Fields = Lines[i].Split(new char[] { ',' });
                Row = dataTable.NewRow();
                for (int f = 0; f < Cols; f++)
                {
                    Row[f] = Fields[f];
                }
                dataTable.Rows.Add(Row);
            }
            return dataTable;
        }
    }
}
