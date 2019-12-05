using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceHander
{
    public partial class Form_CorrectSiteNo : Form
    {
        private Boolean shortSiteNo;
        public Form previousForm;

        public bool ShortSiteNo { get => shortSiteNo; }

        public Form_CorrectSiteNo(Form previousForm)
        {
            InitializeComponent();
            this.previousForm = previousForm;
        }

        private void Button_ok_Click(object sender, EventArgs e)
        {
            if (radioButton_fullName.Checked == true)
                shortSiteNo = false;
            else
                shortSiteNo = true;

            this.Hide();
            this.DialogResult = DialogResult.OK;
            previousForm.Activate();
        }
    }
}
