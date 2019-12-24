using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

        private void signal_payLoad_loaded_successfuly()
        {
            label_statusPayLoad.BackColor = Color.GreenYellow;
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
                CommonOperations
                    .Close_running_workbook(SiGlobalVars.Instance.multiTransWorkbook);

            }
            SiGlobalVars.Instance.multiTransWraps = null;
            SiGlobalVars.Instance.multiTransHeadings = null;
            SiGlobalVars.Instance.multiTransCurrentWorkSheet = null;
            SiGlobalVars.Instance.multiTransWorkbook = null;
            button_step1_AddSiteNO.Enabled = false;
            label_StatusMultiTrans.BackColor = Color.Gray;
        }
        private void clear_payLoad_instance()
        {
            if (SiGlobalVars.Instance.payLoadWrap != null)
            {
                //if excel is open, close it
                if (SiGlobalVars.Instance.multiTransWorkbook != null)
                {
                    CommonOperations
                        .Close_running_workbook(SiGlobalVars.Instance.payLoadWorkbook);

                }
                SiGlobalVars.Instance.payLoadWrap = null;
                SiGlobalVars.Instance.payLoadHeadings = null;
                SiGlobalVars.Instance.payLoadWorkbook = null;
                label_statusPayLoad.BackColor = Color.Gray;

            }
        }
        private void clear_mepStyle_instance()
        {
            if (SiGlobalVars.Instance.mepStyleWorkbook != null)
            {
                //if excel is open, close it

                CommonOperations
                       .Close_running_workbook(SiGlobalVars.Instance.mepStyleWorkbook);
            }
            SiGlobalVars.Instance.mepStyleWraps = null;
            SiGlobalVars.Instance.mepStyleHeadings = null;
            SiGlobalVars.Instance.mepStyleCurrentWorkSheet = null;
            SiGlobalVars.Instance.mepStyleWorkbook = null;
            SiGlobalVars.Instance.mepStyleTimesheetMonthYear = null;
            label_StatusMepSty.BackColor = Color.Gray;
        }
        private void clear_dailyTrans_instance()
        {
            if (SiGlobalVars.Instance.dailyTransWorkbook != null)
            {
                //if excel is open, close it

                CommonOperations
                      .Close_running_workbook(SiGlobalVars.Instance.dailyTransWorkbook);

            }
            SiGlobalVars.Instance.dailyTransWraps = null;
            SiGlobalVars.Instance.dailyTransHeadings = null;
            SiGlobalVars.Instance.dailyTransCurrentWorkSheet = null;
            SiGlobalVars.Instance.dailyTransWorkbook = null;
            label_StatusDailyTrans.BackColor = Color.Gray;
        }

        private void clear_payLoadTrans_instance()
        {
            if (SiGlobalVars.Instance.payLoadWorkbook != null)
            {
                //if excel is open, close it

                CommonOperations
                     .Close_running_workbook(SiGlobalVars.Instance.payLoadWorkbook);

            }
            SiGlobalVars.Instance.payLoadWrap = null;
            SiGlobalVars.Instance.payLoadHeadings = null;
            SiGlobalVars.Instance.payLoadWorkbook = null;
            SiGlobalVars.Instance.Holidays = null;
            label_statusPayLoad.BackColor = Color.Gray;
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

            if (MixTimeSheetHandler
                    .Add_siteNo_from_DailyTrans_to_MultiTrans
                    (SiGlobalVars.Instance.dailyTransWraps,
                    ref SiGlobalVars.Instance.multiTransWraps)
                    == true)
                MessageBox.Show("Successfuly Added Site No.s from " +
                    "Daily Transactions to Multiple Transactions");
            else
                MessageBox.Show("Successfuly Added Site No.s from " +
                           "Daily Transactions to Multiple Transactions");
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
            Form_holidaysSelector form_HolidaysSelector
                 = new Form_holidaysSelector(null);
            form_HolidaysSelector.Show();
        }

        private void Button2_Click(object sender, EventArgs e)
        {

        }


        private void Button_step3_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.multiTransWorkbook == null ||
                SiGlobalVars.Instance.multiTransCurrentWorkSheet == null ||
                SiGlobalVars.Instance.multiTransWraps == null)
            {
                MessageBox.Show("Multiple Trans Workbook & Worksheet are null");
                return;
            }

            String pdf_export_output_path;

            MultipleTransaction.MultiTransHelper multiTransHelper
              = new MultipleTransaction.MultiTransHelper(SiGlobalVars.Instance.multiTransCurrentWorkSheet,
              SiGlobalVars.Instance.multiTransWorkbook);
            if (multiTransHelper.PRINT_each_employee(this,
               out pdf_export_output_path)
                 == true)
                MessageBox.Show("Printing Successfuly Completed. Output Folder = "
                    + pdf_export_output_path);


        }

        private void ButtonTestMultiTrans_Click(object sender, EventArgs e)
        {
            Form_TestMepStyle form_TestMepStyle = new Form_TestMepStyle();
            form_TestMepStyle.Show();

        }

        private void Button_step4_transfDataToMep_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.multiTransWraps == null ||
                SiGlobalVars.Instance.mepStyleWraps == null)
            {
                MessageBox.Show("Either Multiple Transaction or MEP Timesheet is null");
                return;
            }
            MixTimeSheetHandler
                  .Transfer_data_from_multiTrans_to_mepStyle
                  (SiGlobalVars.Instance.multiTransWraps,
                  SiGlobalVars.Instance.mepStyleWraps);



            MessageBox.Show("Data transfer from Multiple Transaciton to Mep Successfuly Completed");
        }

        private void Button_step2A_siteNo_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.multiTransWraps == null)
            {
                MessageBox.Show(" MULTI Trans is null");
                return;
            }
            MixTimeSheetHandler mixTimeSheetHandler = new MixTimeSheetHandler();
            Form_CorrectSiteNo form_CorrectSiteNo = new Form_CorrectSiteNo(this);
            form_CorrectSiteNo.ShowDialog();
            if (form_CorrectSiteNo.DialogResult == DialogResult.OK)
            {
                Boolean shortSiteNo = form_CorrectSiteNo.ShortSiteNo;

                if (shortSiteNo == true)
                {
                    if (mixTimeSheetHandler
                          .correct_siteNo_in_multiTrans_with_shortMechSiteNo
                          (SiGlobalVars.Instance.multiTransWraps) == true)
                        MessageBox.Show("Site No.s are corrected to Short Mech Site No in MultiTrans Excel file");
                    else
                        MessageBox.Show("Site No. correction failed");
                }
                else
                {
                    if (mixTimeSheetHandler
                         .correct_siteNo_in_multiTrans_with_fullMechSiteNo
                         (SiGlobalVars.Instance.multiTransWraps) == true)
                        MessageBox.Show("Site No.s  are corrected to Long Mech Site No in MultiTrans Excel file");
                    else
                        MessageBox.Show("Site No. correction failed");

                }

            }

        }

        private void Button_Step5_MultiTranToPay_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.clearanceFor_step5B_MultiToPay == false)
            {
                MessageBox.Show("Cannot run this task now. This is because " +
                    " Site No should be re-extracted before running this task;" +
                    " Hence the task if failed;");
                return;
            }
            if (SiGlobalVars.Instance.payLoadWrap == null ||
                SiGlobalVars.Instance.multiTransWraps == null)
            {
                MessageBox.Show("Either Multiple Transaction or Payload Transaction is null;" +
                    "Hence Task failed");
                return;
            }
            List<Excel.Workbook> workbooks = new List<Excel.Workbook>()
            {SiGlobalVars.Instance.payLoadWorkbook,
            SiGlobalVars.Instance.multiTransWorkbook};

            if (EXCEL_HELPER.Given_workbooks_are_running_in_background(workbooks) == false)
            {
                MessageBox.Show("Either Multiple Transaction or Daily Transaction" +
                    " Workbooks are not running in background; Hence Task Failed;");
                return;
            }


            if (CommonOperations
                 .Display_holiday_selectorForm(this, label_holidays) == true)
            {
                Boolean printBio_inPayLoad = false;
                DialogResult dialogResult =
                     MessageBox.Show("Print BioWorkTime in Payload?",
                     "Print bio", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                    printBio_inPayLoad = true;

                if (MixTimeSheetHandler
                   .Transfer_data_from_multiTrans_to_payLoad
                   (SiGlobalVars.Instance.Holidays, printBio_inPayLoad)
                     == false)
                    MessageBox.Show("Transfer from Multiple Transaction to PayLoad failed");
                else
                    MessageBox.Show("Succesfully Transferred from Multiple Transaction to PayLoad");
            }
            else
            {
                MessageBox.Show("Holiday Dates are not selected");
                return;
            }





        }





        private void Button_openPayLoad_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = openFile(true);
            if (workbook == null)
                return;

            SiGlobalVars.Instance.payLoadWorkbook = workbook;

            initiate_understanding_payLoad_timeSheet();
            this.Activate();
        }

        private void initiate_understanding_payLoad_timeSheet()
        {

            var workbook = SiGlobalVars.Instance.payLoadWorkbook;
            //TODO: later the headings should be loaded from the settings
            //code for the same should be implemented
            //now the name of the columns are hard coded
            //but later a settings shall be introduced to change the
            //heading names dynamically

            PayLoadFormat.PayLoadHelper payLoadHelper
                = new PayLoadFormat.PayLoadHelper(workbook);
            Boolean error_found = false;

            payLoadHelper.MAIN_understand_the_excel_sheet(out error_found);
            if (error_found == false)
            {
                signal_payLoad_loaded_successfuly();
            }
            else
            {
                clear_payLoad_instance();
            }
        }



        private void Button_step6_transMEPtoPay_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.mepStyleWraps == null ||
                            SiGlobalVars.Instance.payLoadWrap == null)
            {
                MessageBox.Show("MEP Format or Payload Transaction" +
                    " is not found");
                return;
            }

            if (MixTimeSheetHandler.Transfer_MEPdata_to_payLoad() == true)
                MessageBox.Show("Transfer from MEP to PayLoad Successfuly Completed");
            else
                MessageBox.Show("Failed to transfer data from MEP to PayLoad");
        }

        private void button_clearPayLoad_Click(object sender, EventArgs e)
        {
            clear_payLoadTrans_instance();
        }

        private void Button_Step5A_Click(object sender, EventArgs e)
        {
            if (SiGlobalVars.Instance.multiTransWraps == null)
            {
                MessageBox.Show("Multiple Transactions is null");
                return;
            }


            if (MultipleTransaction.MultiTransHelper
                  .ReExtract_siteNos_from_multipleTransactions_cells() == true)
            {

                if (MultipleTransaction.MultiTransHelper
                           .AutoFill_SiteNos_for_fridaysAndHolidays() == true)
                {
                    MessageBox.Show("The Task Auto Fill Site No.s for fridays & Holidays Successfully Completed ");
                    SiGlobalVars.Instance.clearanceFor_step5B_MultiToPay = true;
                    //after running the autofill
                    //again we have to run re-extract site no.s
                    MultipleTransaction.MultiTransHelper
                  .ReExtract_siteNos_from_multipleTransactions_cells();
                    MessageBox.Show("Successfully Re-extracted Site No.s from the cells of Multiple Transactions");


                }
                else
                {
                    MessageBox.Show("Failed to complete the Task Auto Fill Site No.s for fridays & Holidays");
                }


            }
            else
                MessageBox.Show("Failed to Re-extract site no.s from the cells of Mulitple " +
                    "Transactions");


        }
    }
}
