using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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

        private Excel.Workbook openFile()
        {
            InputOutHandler inputOutHandler =
               new InputOutHandler(openFileDialog1);
            FileInfo file = inputOutHandler.open_file();
            String filename = file.FullName;
            if (file == null)
                return null;

            Excel.Application application = Globals.ThisAddIn.getApplication();
            return (application.Workbooks.Open(filename));


        }

        private void ButtonOpenMepStyle_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = openFile();
            if (workbook == null)
                return;

            SI_GlobalVars.Instance.mepStyleTimeSheet = workbook;
            FormSheetSelector form = new FormSheetSelector(workbook,
                this);
            form.Show();

        }

        private void ButtonTestMepStyle_Click(object sender, EventArgs e)
        {
            TestMepStyleTime testMepStyleTime
                = new TestMepStyleTime();
            this.Hide();
            testMepStyleTime.Show();
        }
    }
}
