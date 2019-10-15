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

        private void open_file(IfileOpenAction ifileOpenAction)
        {
            DialogResult dialogResult;
            List<String> filenames = new List<String>();
            dialogResult = openFileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                filenames = openFileDialog.FileNames.ToList<String>();
                List<FileInfo> files = get_file_from_filename_strings(filenames);
                ifileOpenAction.okButtonPressed(files);
            }
            else
            {
                ifileOpenAction.cancelled();
            }

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
        private void open_files(IfileOpenAction ifileOpenAction)
        {
            DialogResult dialogResult;
            dialogResult = openFileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                String filename = openFileDialog.FileName;
                FileInfo file = new FileInfo(filename);
                ifileOpenAction.okButtonPressed(file);
            }
            else
            {
                ifileOpenAction.cancelled();
            }

        }
    }
}
