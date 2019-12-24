using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AttendanceHander.MultipleTransaction;

namespace AttendanceHander
{
    public partial class Form_AutoFillDialog : Form
    {
        private Form previousForm;
        private Boolean autoFillSiteNo;
        public Form_AutoFillDialog()
        {
            InitializeComponent();
        }

        public Form_AutoFillDialog(Form previousForm)
        {
            this.previousForm = previousForm;
        }

        public bool AutoFillSiteNo { get => autoFillSiteNo;  }

        private void Button_proceed_Click(object sender, EventArgs e)
        {
            autoFillSiteNo = checkBox_autoFillSite.Checked;
            this.Hide();
            this.DialogResult = DialogResult.OK;
            previousForm.Activate();
        }
    }
}
