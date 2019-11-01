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

        public void enableDisableMepStyleTestButton(Boolean enable)
        {
            buttonTestMepStyle.Enabled = enable;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Worksheet current_worksheet = Globals.ThisAddIn.get_active_worksheet();
            AttendHelper attendHelper = new AttendHelper();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {

        }

        private Excel.Workbook openFile(Boolean readOnly)
        {
            InputOutHandler inputOutHandler =
               new InputOutHandler(openFileDialog1);
            FileInfo file = inputOutHandler.open_file();
            if (file == null)
                return null;
            String filename = file.FullName;
            if (file == null)
                return null;

            Excel.Application application = Globals.ThisAddIn.getApplication();
            return (application.Workbooks.Open(filename, ReadOnly: readOnly));


        }
        private void initiate_understanding_MultipleTransaction_timesheet()
        {

            var workbook = SiGlobalVars.Instance.multiTransWorkbook;
            var worksheet = SiGlobalVars.Instance.multiTransCurrentWorkSheet;
            //TODO: later the headings should be loaded from the settings
            //code for the same should be implemented
            //now the name of the columns are hard coded
            //but later a settings shall be introduced to change the
            //heading names dynamically

            MultipleTransaction.MultiTransHelper multiTransHelper
                = new MultipleTransaction.MultiTransHelper(worksheet, workbook);

            multiTransHelper.MAIN_understand_the_excel_sheet();
        }

        private void initiate_understanding_dailyTrans_timesheet()
        {

            var workbook = SiGlobalVars.Instance.dailyTransWorkbook;
            var worksheet = SiGlobalVars.Instance.dailyTransCurrentWorkSheet;
            //TODO: later the headings should be loaded from the settings
            //code for the same should be implemented
            //now the name of the columns are hard coded
            //but later a settings shall be introduced to change the
            //heading names dynamically
            DailyTransactions.DailyTransHelper dailyTransHelper
                = new DailyTransactions.DailyTransHelper(worksheet, workbook);


            dailyTransHelper.MAIN_understand_the_excel_sheet();
        }

        private void initiate_understanding_MEP_style_timesheet()
        {

            var workbook = SiGlobalVars.Instance.mepStyleWorkbook;
            var worksheet = SiGlobalVars.Instance.mepStyleCurrentWorkSheet;
            //TODO: later the headings should be loaded from the settings
            //code for the same should be implemented
            //now the name of the columns are hard coded
            //but later a settings shall be introduced to change the
            //heading names dynamically

            MepStyleHelper mepStyleHelper =
                new MepStyleHelper(workbook,
                worksheet);
            mepStyleHelper.MAIN_understand_the_excel_sheet();
        }

        private void ButtonTestMepStyle_Click(object sender, EventArgs e)
        {
            TestMepStyleTime testMepStyleTime
                = new TestMepStyleTime();
            this.Hide();
            testMepStyleTime.Show();
        }

        private void FormMain_Activated(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.mepStyleCurrentWorkSheet != null)
                buttonTestMepStyle.Enabled = true;
        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }

        
        private void buttonOpenMultiTrans_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = openFile(true);
            if (workbook == null)
                return;

            SiGlobalVars.Instance.multiTransWorkbook = workbook;
            FormSheetSelector form = new FormSheetSelector(workbook,
               (FormMain)this);
            form.ShowDialog();
            Excel.Worksheet selected_Sheet = form.Selected_sheet;
            if (selected_Sheet == null)
            {
                MessageBox.Show("Selected Sheet is null");
                return;
            }
            else
            {
                SiGlobalVars.Instance.multiTransCurrentWorkSheet = selected_Sheet;
            }

            if (form.DialogResult == DialogResult.OK)
            {
                initiate_understanding_MultipleTransaction_timesheet();
                buttonTestMultiTrans.Enabled = true;
                this.Activate();
            }
        }

        private void ButtonOpenMepTimesheet_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = openFile(true);
            if (workbook == null)
                return;

            SiGlobalVars.Instance.mepStyleWorkbook = workbook;
            FormSheetSelector form = new FormSheetSelector(workbook,
               (FormMain)this);
            form.ShowDialog();
            Excel.Worksheet selected_Sheet = form.Selected_sheet;
            if (selected_Sheet == null)
            {
                MessageBox.Show("Selected Sheet is null");
                return;
            }
            else
            {
                SiGlobalVars.Instance.mepStyleCurrentWorkSheet = selected_Sheet;
            }

            if (form.DialogResult == DialogResult.OK)
            {
                initiate_understanding_MEP_style_timesheet();
                buttonTestMepStyle.Enabled = true;

                this.Activate();
            }

        }

        private void Button_OpenDailyTrans_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = openFile(true);
            if (workbook == null)
                return;

            SiGlobalVars.Instance.dailyTransWorkbook = workbook;
            FormSheetSelector form = new FormSheetSelector(workbook,
               (FormMain)this);
            form.ShowDialog();
            Excel.Worksheet selected_Sheet = form.Selected_sheet;
            if (selected_Sheet == null)
            {
                MessageBox.Show("Selected Sheet is null");
                return;
            }
            else
            {
                SiGlobalVars.Instance.dailyTransCurrentWorkSheet = selected_Sheet;
            }

            if (form.DialogResult == DialogResult.OK)
            {
                initiate_understanding_dailyTrans_timesheet();
                button_OpenDailyTrans.Enabled = true;

                this.Activate();
            }
        }

        private void Button_step1_AddSiteNO_Click(object sender, EventArgs e)
        {
            if(SiGlobalVars.Instance.multiTransWraps==null ||
                SiGlobalVars.Instance.dailyTransWraps==null)
            {
                MessageBox.Show("Multiple Transaction or Daily Transaction is null");
                return;
            }

            MixTimeSheetHandler
                .Add_siteNo_from_DailyTrans_to_MultiTrans
                (SiGlobalVars.Instance.dailyTransWraps,
                ref SiGlobalVars.Instance.multiTransWraps);

        }

        private void Button_TestDailyTrans_Click(object sender, EventArgs e)
        {

        }
    }
}
