using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace AttendanceHander
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Worksheet current_worksheet = Globals.ThisAddIn.get_active_worksheet();
            AttendHelper attendHelper = new AttendHelper();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();

            Excel.Worksheet current_worksheet = Globals.ThisAddIn.get_active_worksheet();
            label1.Text = current_worksheet.Name;
            AttendHelper attendHelper = new AttendHelper();
        }
    }
}
