using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceHander
{
    class InputOutHandler
    {
        OpenFileDialog openFileDialog;

        public InputOutHandler(OpenFileDialog openFileDialog)
        {
            this.openFileDialog = openFileDialog;
        }

        public List<FileInfo> open_files()
        {
            DialogResult dialogResult;
            List<String> filenames = new List<String>();
            dialogResult = openFileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                filenames = openFileDialog.FileNames.ToList<String>();
                List<FileInfo> files = get_file_from_filename_strings(filenames);
                return files;
            }
            return null;

        }

        private List<FileInfo> get_file_from_filename_strings(List<String> filenames)
        {
            List<FileInfo> files = new List<FileInfo>();
            foreach (String filename in filenames)
            {
                FileInfo file = new FileInfo(filename);
                files.Add(file);
            }
            return files;
        }
        public FileInfo open_file()
        {
            DialogResult dialogResult;
            dialogResult = openFileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                String filename = openFileDialog.FileName;
                FileInfo file = new FileInfo(filename);
                return file;
            }
            return null;

        }
    }
}
