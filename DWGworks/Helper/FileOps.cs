using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace TunnalCal.Helper
{
    class FileOps
    {
        public static string SelectFile()
        {
            //abtain save file path
            OpenFileDialog ODialog = new OpenFileDialog();
            //FolderBrowserDialog ODialog = new FolderBrowserDialog();
            string fileFullame = "";
            if (ODialog.ShowDialog() == DialogResult.OK)
            {
                fileFullame = ODialog.FileName;
            }

            return fileFullame;
        }

        public static string[] selectFiles(string title, string filter, string dialog)
        {
            //use autodesk windows rather than windows form
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog(title, null, filter, dialog, Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            ofd.ShowDialog();
            string[] result = ofd.GetFilenames();

            return result;
        }

        public static string[] getFilePaths(string messageTitle, string filter, bool selectMultiFiles)
        {
            string[] filePath = null;
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();

                string CurrentDirectory = Environment.CurrentDirectory;//set current directory to be environment's current directory
                ofd.RestoreDirectory = false;//set RestoreDirectory to false
                string selectedDirectory;
                ofd.Title = messageTitle;
                ofd.Filter = filter;
                ofd.Multiselect = selectMultiFiles;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    selectedDirectory = Environment.CurrentDirectory;
                    ofd.InitialDirectory = selectedDirectory;

                    filePath = ofd.FileNames;

                    Environment.CurrentDirectory = CurrentDirectory;
                }
            }
            catch (Exception ex)
            {
            }

            return filePath;
        }

    }
}
