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


        private void signal_multiTrans_loaded_successfuly()
        {
            button_step1_AddSiteNO.Enabled = true;
            buttonTestMultiTrans.Enabled = true;
            label_StatusMultiTrans.BackColor = Color.GreenYellow;
        }
        private void signal_dailyTrans_loaded_successfuly()
        {
            button_step1_AddSiteNO.Enabled = true;
            button_TestDailyTrans.Enabled = true;
            label_StatusDailyTrans.BackColor = Color.GreenYellow;
        }
        private void signal_mepStyle_loaded_successfuly()
        {
            buttonTestMepStyle.Enabled = true;
            label_StatusMepSty.BackColor = Color.GreenYellow;
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

        private void clear_multiTrans_instance()
        {
            if (SiGlobalVars.Instance.multiTransWorkbook != null)
            {
                //if excel is open, close it
                SiGlobalVars.Instance.multiTransWorkbook.Close();
                SiGlobalVars.Instance.multiTransWraps = null;
                SiGlobalVars.Instance.multiTransHeadings = null;
                SiGlobalVars.Instance.multiTransCurrentWorkSheet = null;
                SiGlobalVars.Instance.multiTransWorkbook = null;
                button_step1_AddSiteNO.Enabled = false;
                label_StatusMultiTrans.BackColor = Color.Gray;

            }
        }

        private void clear_mepStyle_instance()
        {
            if (SiGlobalVars.Instance.mepStyleWorkbook != null)
            {
                //if excel is open, close it
                SiGlobalVars.Instance.mepStyleWorkbook.Close();
                SiGlobalVars.Instance.mepStyleWraps = null;
                SiGlobalVars.Instance.mepStyleHeadings = null;
                SiGlobalVars.Instance.mepStyleCurrentWorkSheet = null;
                SiGlobalVars.Instance.mepStyleWorkbook = null;
                label_StatusMepSty.BackColor = Color.Gray;

            }
        }
        private void clear_dailyTrans_instance()
        {
            if (SiGlobalVars.Instance.dailyTransWorkbook != null)
            {
                //if excel is open, close it
                SiGlobalVars.Instance.dailyTransWorkbook.Close();
                SiGlobalVars.Instance.dailyTransWraps = null;
                SiGlobalVars.Instance.dailyTransHeadings = null;
                SiGlobalVars.Instance.dailyTransCurrentWorkSheet = null;
                SiGlobalVars.Instance.dailyTransWorkbook = null;
                label_StatusDailyTrans.BackColor = Color.Gray;

            }
        }



        private void initiate_understanding_MultipleTransaction_timesheet()
        {
            button_step1_AddSiteNO.Enabled = false;

            var workbook = SiGlobalVars.Instance.multiTransWorkbook;
            var worksheet = SiGlobalVars.Instance.multiTransCurrentWorkSheet;
            //TODO: later the headings should be loaded from the settings
            //code for the same should be implemented
            //now the name of the columns are hard coded
            //but later a settings shall be introduced to change the
            //heading names dynamically

            MultipleTransaction.MultiTransHelper multiTransHelper
                = new MultipleTransaction.MultiTransHelper(worksheet, workbook);
            Boolean error_found = false;
            multiTransHelper.MAIN_understand_the_excel_sheet(out error_found);
            if (error_found == false)
            {
                signal_multiTrans_loaded_successfuly();
            }
            else
            {
                clear_multiTrans_instance();
            }
        }

        private void initiate_understanding_dailyTrans_timesheet()
        {
            button_step1_AddSiteNO.Enabled = false;

            var workbook = SiGlobalVars.Instance.dailyTransWorkbook;
            var worksheet = SiGlobalVars.Instance.dailyTransCurrentWorkSheet;
            //TODO: later the headings should be loaded from the settings
            //code for the same should be implemented
            //now the name of the columns are hard coded
            //but later a settings shall be introduced to change the
            //heading names dynamically
            DailyTransactions.DailyTransHelper dailyTransHelper
                = new DailyTransactions.DailyTransHelper(worksheet, workbook);

            Boolean error_found = false;
            dailyTransHelper.MAIN_understand_the_excel_sheet(out error_found);

            if (error_found == true)
            {
                clear_dailyTrans_instance();
            }
            else
            {
                button_step1_AddSiteNO.Enabled = true;
                signal_dailyTrans_loaded_successfuly();
            }


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

            Boolean error_found = false;
            mepStyleHelper.MAIN_understand_the_excel_sheet(out error_found);
            if (error_found == true)
                clear_mepStyle_instance();
            else
                signal_mepStyle_loaded_successfuly();
        }

        private void ButtonTestMepStyle_Click(object sender, EventArgs e)
        {
            Form_TestMepStyle testMepStyleTime
                = new Form_TestMepStyle();
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
            if (SiGlobalVars.Instance.multiTransWraps == null ||
                SiGlobalVars.Instance.dailyTransWraps == null)
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

        private void Button_step2_missingData_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.multiTransWraps == null ||
                SiGlobalVars.Instance.mepStyleWraps == null)
            {
                MessageBox.Show("Either MEP STYLE or MULTI Trans is null");
                return;
            }
            MixTimeSheetHandler mixTimeSheetHandler = new MixTimeSheetHandler();
            mixTimeSheetHandler
                .Add_Missing_data_from_mepStyle_to_MultiTrans(SiGlobalVars.Instance.mepStyleWraps,
               ref SiGlobalVars.Instance.multiTransWraps);
        }

        private void Button_clearMultiTrans_Click(object sender, EventArgs e)
        {
            clear_multiTrans_instance();
        }

        private void Button_clearMepStyle_Click(object sender, EventArgs e)
        {
            clear_mepStyle_instance();
        }

        private void Button_clearDailyTrans_Click(object sender, EventArgs e)
        {
            clear_dailyTrans_instance();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            String simply = "just to break the program";
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            
        }

        private void Button_step3_Click(object sender, EventArgs e)
        {
            if(SiGlobalVars.Instance.multiTransWorkbook==null||
                SiGlobalVars.Instance.multiTransCurrentWorkSheet==null||
                SiGlobalVars.Instance.multiTransWraps==null)
            {
                MessageBox.Show("Multiple Trans Workbook & Worksheet are null");
                return;
            }
            MultipleTransaction.MultiTransHelper multiTransHelper
                = new MultipleTransaction.MultiTransHelper(SiGlobalVars.Instance.multiTransCurrentWorkSheet,
                SiGlobalVars.Instance.multiTransWorkbook);
            if (multiTransHelper.PRINT_each_employee(folderBrowserDialog_PDFexport, this)
                  == true)
                MessageBox.Show("Printing Successfuly Completed. Output Folder = "
                    + folderBrowserDialog_PDFexport.SelectedPath);


        }

        private void ButtonTestMultiTrans_Click(object sender, EventArgs e)
        {
            Form_TestMepStyle form_TestMepStyle = new Form_TestMepStyle();
            form_TestMepStyle.Show();

        }
    }
}
