using System;
using System.Windows.Forms;

namespace ExPat_UI.Utilities
{
    class FileUtilities
    {
        /// <summary> Public Method PickFolder
        /// Accepts null or default start location for dialog box.
        /// </summary>
        /// <param name="defaultLocation">null or default location to open dialog</param>
        /// <returns>Folder Path</returns>
        public string PickFolder(string defaultLocation)
        {
            string folderName = null;
            FolderBrowserDialog selectedFolder = new FolderBrowserDialog();
            if (defaultLocation == null)
            {
                defaultLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            //selectedFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            selectedFolder.SelectedPath = defaultLocation;
            DialogResult result = selectedFolder.ShowDialog();
            if (result == DialogResult.OK)
            {
                folderName = selectedFolder.SelectedPath.Replace(@"\\", @"\") + @"\";
            }
            else
            {
                throw new System.IO.FileNotFoundException("Folder Path was not selected.");
            }
            return folderName;
        }

        /// <summary> Public Method OpenFile
        /// Accepts  null or default location and fileType (to be expanded) to be opened.
        /// </summary>
        /// <param name="defaultLocation">null or default location to open dialog.</param>
        /// <param name="fileType">filetype to look for.</param>
        /// <returns>Path and Filename of selected file</returns>
        public string OpenFile(string defaultLocation, string fileType, string title)
        {
            string fileName = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Text files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog.Title = title.ToString();
            if (defaultLocation == null)
            {
                defaultLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            openFileDialog.InitialDirectory = defaultLocation;

            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
            }
            else
            {
                throw new System.IO.FileNotFoundException("Import File was not selected.");
            }
            return fileName;
        }
    }
}
