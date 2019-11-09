using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceHander
{
    public partial class Form_ExportStatus : Form
    {
        Thread exportingThread;
        public Form_ExportStatus(ref Thread exportingThread)
        {
            InitializeComponent();
            this.exportingThread = exportingThread;
        }

        private void Button_cancelExport_Click(object sender, EventArgs e)
        {
            
        }

        private void Form_ExportStatus_Load(object sender, EventArgs e)
        {

        }
    }
}
