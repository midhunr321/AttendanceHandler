using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander.MultipleTransaction
{
    public class MultiTransHelper
    {
        private Excel.Worksheet worksheet;
        private Excel.Workbook workbook;

        public MultiTransHelper(Excel.Worksheet worksheet, Excel.Workbook workbook)
        {
            this.worksheet = worksheet;
            this.workbook = workbook;
        }

        public class MultiHeadings : IEnumerable<HeadingWrap>
        {
            public HeadingWrap sheetHeading =
                new HeadingWrap("Multiple Transaction");
            public HeadingWrap personnelNo = new HeadingWrap("Personnel No.");
            public HeadingWrap firstName = new HeadingWrap("First Name");
            public HeadingWrap lastName = new HeadingWrap("Last Name");
            public HeadingWrap position = new HeadingWrap("Position");
            public HeadingWrap department = new HeadingWrap("Department");
            public HeadingWrap date = new HeadingWrap("Date");
            public HeadingWrap checkInTime1 = new HeadingWrap("Check-In");
            public HeadingWrap checkOutTime1 = new HeadingWrap("Check-Out");
            public HeadingWrap workingTime1 = new HeadingWrap("Working Time");
            public HeadingWrap checkInTime2 = new HeadingWrap("Check-In");
            public HeadingWrap checkOutTime2 = new HeadingWrap("Check-Out");
            public HeadingWrap workingTime2 = new HeadingWrap("Working Time");
            public HeadingWrap totalTimeWorked = new HeadingWrap("Total Time Worked");
            public HeadingWrap siteNo; //this shouldn't be iteration. because siteno doesn't exist initially
            //we need to create it in future.
            public Dictionary<int, HeadingWrap> overtimeDays;

            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {sheetHeading,personnelNo,firstName,lastName
                ,position,department,date,checkInTime1,checkOutTime1,workingTime1,
                checkInTime2,checkOutTime2,workingTime2,
                totalTimeWorked}.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }


        }


        public static Boolean Add_a_heading_column_for_site_no(Excel.Range previousHeading,
            Excel.Worksheet sheet, ref MultiHeadings headingWraps)
        {
            if (sheet == null)
            {
                MessageBox.Show("Multiple Transaction Excel file is either closed or invalid");
                return false;
            }

            Excel.Range siteNoHeading
                = previousHeading.Next;

            if (headingWraps.siteNo != null)
            {
                return true;
            }
            else
            {
                headingWraps.siteNo = new HeadingWrap("Site No.");
                headingWraps.siteNo.fullCell = siteNoHeading;

                //now create the new heading in the excel worksheet
                headingWraps.siteNo.fullCell.Value = "Site No.";


                return true;
            }

        }

        private void hide_unhide_all_multi_trans_data_rows(Boolean hide)
        {
            foreach (var datawrap in SiGlobalVars.Instance.multiTransWraps)
            {
                datawrap.personnelNo.fullCell.EntireRow.Hidden = hide;
            }
        }

        private void hide_unhide_data_of(IGrouping<String, MultiTransWrap> igroup, Boolean hide,
            out String employeeName)
        {
            employeeName = null;
            foreach (var item in igroup)
            {
                item.personnelNo.fullCell.EntireRow.Hidden = hide;
                employeeName = item.firstName.content;
            }
        }

        private FileInfo get_output_file(DirectoryInfo outputDirectory, String filename_without_extension)
        {


            if (outputDirectory == null)
                return null;

            String outputPath = outputDirectory.FullName + "\\" + filename_without_extension
                   + ".pdf";

            FileInfo outputFile = new FileInfo(outputPath);

            return outputFile;
        }
        private DirectoryInfo open_directory_dialog_for_exporting_PDF()
        {
            DirectoryInfo folder = null;
            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();
            commonOpenFileDialog.IsFolderPicker = true;
            if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {

                folder = new DirectoryInfo(commonOpenFileDialog.FileName);
            }

            return folder;
            //DialogResult dialogResult = folderBrowserDialog.ShowDialog();
            //DirectoryInfo directoryInfo;
            //if (dialogResult == DialogResult.OK)
            //{
            //    directoryInfo = new DirectoryInfo(folderBrowserDialog.SelectedPath);
            //    return directoryInfo;
            //}
            //else
            //{
            //    return null;
            //}

        }

        private String select_employeePosition(Form rootForm)
        {
            //Todo: Employee positions are explicity defined instead of finding from the worksheet


            Selector selector = new Selector("Employee Position",
                SiGlobalVars.Instance.assumed_MultiTrans_EmployeePositions,
               rootForm);
            selector.ShowDialog();
            String selected_employee_position;
            DialogResult dialogResult = selector.DialogResult;
            if (dialogResult == DialogResult.OK)
            {
                selected_employee_position = selector.selected_employeePosition;

                return selected_employee_position;

            }

            return null;

        }
        public Boolean PRINT_each_employee(Form main_form, out String outputPath)
        {
            outputPath = null;
            //first hide all cells
            hide_unhide_all_multi_trans_data_rows(true);

            //now show rows one by one for each employee
            List<MultiTransWrap> multiTransWraps = SiGlobalVars.Instance.multiTransWraps;
            String selected_employeePosition = select_employeePosition(main_form);
            if (selected_employeePosition == null)
                return false;

            var filter_by_position = multiTransWraps
                .Where(x => x.position.content == selected_employeePosition).ToList();

            var group_by_employeeNo = filter_by_position.GroupBy(x => x.personnelNo.content).ToList();

            DirectoryInfo outputDirectory =
                open_directory_dialog_for_exporting_PDF();

            if (outputDirectory == null)
            {
                MessageBox.Show("Output Directory wasn't selected");
                return false;
            }

            foreach (var igroup in group_by_employeeNo)
            {
                String employeeName;
                //first  show the selected rows
                hide_unhide_data_of(igroup, false, out employeeName);
                //now print the single employee rows

                String fileName_without_extension = igroup.Key + " " + employeeName;
                FileInfo outputFile = get_output_file(outputDirectory,
                    fileName_without_extension);
                EXCEL_HELPER.print_to_pdf(SiGlobalVars.Instance.multiTransCurrentWorkSheet,
                    outputFile);

                //after exporting the pdf, again we have to hide the rows of the ones we have
                //already exported as pdf. So..
                hide_unhide_data_of(igroup, true, out employeeName);

            }

            outputPath = outputDirectory.FullName;

            return true;
        }


        public void MAIN_understand_the_excel_sheet(out Boolean error_found)
        {
            if (SiGlobalVars.Instance.multiTransHeadings == null)
            {
                SiGlobalVars.Instance.multiTransHeadings
                    = new MultiTransHelper.MultiHeadings();
            }
            if (SiGlobalVars.Instance.multiTransWraps == null)
                SiGlobalVars.Instance.multiTransWraps = new List<MultiTransWrap>();

            find_headings(ref SiGlobalVars.
                Instance.multiTransHeadings, out error_found);
            if (error_found == true)
                return;


            //now we got all the headings
            //connect heading and datas together

            //now that we got all headings
            //we need to start with the rows

            read_each_rows_of_data(out error_found);
            if (error_found == true)
                return;

        }
        private Boolean find_headings(ref MultiTransHelper
            .MultiHeadings headingWraps,
            out Boolean error_found)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            foreach (HeadingWrap heading in headingWraps)
            {
                //sometimes for special cases like check-in time 1 and check in time 2
                //when we search for keyword "check-in" both results are obtained at the same time
                //so at the same iteration instance we set full cell for these both.
                //so in such cases just continue the iteration without enterring for check in time 2
                if (heading.fullCell != null)
                    continue;


                List<Excel.Range> search_results = new List<Excel.Range>();
                search_results =
                    eXCEL_HELPER.find_fix_column_heading(heading.headingName,
                    Excel.XlSearchDirection.xlNext,
                    Excel.XlSearchOrder.xlByRows, false);

                if (search_results == null)
                {
                    MessageBox.Show("Couldn't find the heading = "
                        + heading.headingName);
                    error_found = true;
                    return false;
                }
                //TODO: Should carryout the search from the top

                // if the search count is not more than 1 then,
                if (search_results != null && search_results.Count == 1)
                {

                    check_if_special_case(search_results, heading,
                        out error_found);
                    if (error_found == true)
                        return false;

                    Excel.Range fullcell = eXCEL_HELPER
                        .return_full_merg_cell(search_results[0]);
                    heading.fullCell = fullcell;
                }
                else
                {
                    //TODO: if more than one search results
                    //we need to filter it out
                    //like check if the full cell is within the same heading row
                    //that way we can filter out other results.
                    if (check_if_special_case(search_results, heading,
                        out error_found)
                       == true)
                    {
                        //if more than 1 search results means it should be special case
                        if (filterout_multiple_search_results_for_special_case_and_assign_values
                          (search_results, heading)
                              == false)
                        {
                            error_found = true;
                            return false;
                        }


                    }

                    else
                    {
                        Excel.Range assumed_adjacent_cell1 = headingWraps.firstName.fullCell;
                        Excel.Range assumed_adjacent_cell2 = headingWraps.personnelNo.fullCell;


                        Excel.Range filtered_search_result
                            = CommonOperations
                            .filter_searchResult_by_comparing_row_no_of_adjacent_headings
                            (search_results, assumed_adjacent_cell1,
                            assumed_adjacent_cell1);

                        if (filtered_search_result != null)
                        {
                            Excel.Range fullcell = eXCEL_HELPER
                       .return_full_merg_cell(filtered_search_result);
                            heading.fullCell = fullcell;
                        }
                        else
                        {
                            error_found = true;
                            return false;
                        }

                    }
                }

            }



            return true;
        }



        private Boolean filterout_multiple_search_results_for_special_case_and_assign_values
            (List<Excel.Range> searchResults, HeadingWrap heading)
        {

            if (searchResults.Count != 2)
            {
                MessageBox.Show("search results for" +
                    " heading name = " + heading.headingName + " must be 2 no.s");
                return false;
            }

            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            if (eXCEL_HELPER.cells_are_in_the_same_row(searchResults)
                            == false)
                return false;


            var multiTransHeading = SiGlobalVars.Instance.multiTransHeadings;

            if (heading.Equals(multiTransHeading.checkInTime1))
            {

                //lowest column no is obviously check in time 1
                Excel.Range checkInTime1 =
                    eXCEL_HELPER.get_lowest_column_cell_from_search_result(searchResults);

                //and the highest column no is check in time 2 ; so..
                Excel.Range checkIn_time2 =
                 eXCEL_HELPER.get_largest_column_cell_from_search_result(searchResults);

                SiGlobalVars.Instance.multiTransHeadings.checkInTime1.fullCell = checkInTime1;
                SiGlobalVars.Instance.multiTransHeadings.checkInTime1.headingName
                    = eXCEL_HELPER.get_value_of_merge_cell(checkInTime1);

                SiGlobalVars.Instance.multiTransHeadings.checkInTime2.fullCell = checkIn_time2;
                SiGlobalVars.Instance.multiTransHeadings.checkInTime2.headingName
                    = eXCEL_HELPER.get_value_of_merge_cell(checkIn_time2);

            }
            else if (heading.Equals(multiTransHeading.checkOutTime1))
            {

                Excel.Range checkout_time1 =
                    eXCEL_HELPER.get_lowest_column_cell_from_search_result(searchResults);

                Excel.Range checkout_time2 =
                 eXCEL_HELPER.get_largest_column_cell_from_search_result(searchResults);

                SiGlobalVars.Instance.multiTransHeadings.checkOutTime1.fullCell = checkout_time1;
                SiGlobalVars.Instance.multiTransHeadings.checkOutTime1.headingName
                    = eXCEL_HELPER.get_value_of_merge_cell(checkout_time1);

                SiGlobalVars.Instance.multiTransHeadings.checkOutTime2.fullCell = checkout_time2;
                SiGlobalVars.Instance.multiTransHeadings.checkOutTime2.headingName
                    = eXCEL_HELPER.get_value_of_merge_cell(checkout_time2);

            }
            else if (heading.Equals(multiTransHeading.workingTime1))
            {

                Excel.Range workTime1 =
                    eXCEL_HELPER.get_lowest_column_cell_from_search_result(searchResults);

                Excel.Range workTime2 =
                 eXCEL_HELPER.get_largest_column_cell_from_search_result(searchResults);

                SiGlobalVars.Instance.multiTransHeadings.workingTime1.fullCell = workTime1;
                SiGlobalVars.Instance.multiTransHeadings.workingTime1.headingName
                    = eXCEL_HELPER.get_value_of_merge_cell(workTime1);

                SiGlobalVars.Instance.multiTransHeadings.workingTime2.fullCell = workTime2;
                SiGlobalVars.Instance.multiTransHeadings.workingTime2.headingName
                    = eXCEL_HELPER.get_value_of_merge_cell(workTime2);

            }
            return true;
        }

        private Boolean check_if_special_case(List<Excel.Range> search_results,
            HeadingWrap heading, out Boolean error_found)
        {
            //special case means
            //there is two check in , check out & work time
            //we need to check which is what
            //for example checkin_time1 should be before checkin_time2 
            //so does chectout_time1 ...and like that

            //so what we need to check is that
            // since checkin, checkout etc are there 2 times
            //always when we search we need to get two search results
            //if we don't get two search results
            //it mean something is wrong
            error_found = false;

            var multiTransHeading = SiGlobalVars.Instance.multiTransHeadings;

            if (heading.Equals(multiTransHeading.checkInTime1) ||
               heading.Equals(multiTransHeading.checkOutTime1) ||
               heading.Equals(multiTransHeading.workingTime1) ||
               heading.Equals(multiTransHeading.checkInTime2) ||
               heading.Equals(multiTransHeading.checkOutTime2) ||
               heading.Equals(multiTransHeading.workingTime2)
                )
            {
                //means special case;
                //that means there should be two search results
                if (search_results.Count < 2)
                {
                    error_found = true;
                    //if only one search result then it means error is there

                }
                else
                {
                    return true;
                }


            }
            else
            {
                //means it is not special case
                //so return false to indicate that it is not special case
                return false;
            }


            if (error_found == true)
            {
                MessageBox.Show("There should be more than 1 search result for this heading but" +
                    " only 1 search result was found. Detail: Cell Adress = " +
                    search_results[0].Address.ToString());
                return true;
            }

            return false;
        }

        private void read_each_rows_of_data(out Boolean error_occured)
        {
            error_occured = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            //now we have to read the rows
            //our beginning position to start reading the rows will be
            //below the personal Number heading cell

            Excel.Range personnelNo = SiGlobalVars.Instance
                .multiTransHeadings.personnelNo.fullCell;


            //now go to below adjacent cell to personnal no.
            Excel.Range firstDataRowCell =
                eXCEL_HELPER
                .return_immediate_below_cell(personnelNo);



            foreach (Excel.Range row in worksheet.UsedRange.Rows)
            {
                int currentRow = row.Row;
                //our first data row starts from firstDataRowCell
                //so skip the rows above (which are headings)
                if (currentRow < firstDataRowCell.Row)
                    continue;

                //read row will fail at the end of time sheet
                //we need to stop the iteration after plumber no 50
                //after that it is just empty space
                //so if reached_empty_space_area = true, we came accross the empty space
                //thus this iteration can be stopped.

                Boolean reached_empty_space_area = false;
                read_row(row,
                    ref SiGlobalVars.Instance.multiTransWraps,
                    out error_occured, out reached_empty_space_area);
                if (error_occured == true)
                    return;
                if (reached_empty_space_area == true)
                    break;

            }

        }

        private void read_row(Excel.Range row,
         ref List<MultiTransWrap> multiTransWraps,
        out Boolean error_occured,
        out Boolean reached_empty_space_area)
        {
            reached_empty_space_area = false;
            error_occured = false;
            //one way to identify whether we are in empty space
            //that means whether we already passed 50 numbers of plumbers
            //is by detecting if serial no,employee no and name etc are empty
            //if the serial no, employee no and name is empty means 
            //we have reached the end of the time sheet
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            int rowIndex = row.Row;

            Excel.Range firstCell = worksheet.Cells[rowIndex, 1];
            Excel.Range firstFullCell = eXCEL_HELPER.return_full_merg_cell(firstCell);
            Excel.Range nextCell = null;

            Excel.Range nextFullCell = firstFullCell;
            int totalNoUsedColumns = worksheet.UsedRange.Columns.Count;

            int i = 1;
            //to check if we have reached the empty space or blank area after 
            // if no name or employee no is found in this row
            // then we can say we reached the empty space

            MultiTransWrap multiTransWrap = new MultiTransWrap();
            do
            {
                //first nextFullCell is firstFullCell
                //so 
                var currentFullCell = nextFullCell;

                Boolean result1;
                result1 = feed_datas_of_single_row(ref multiTransWrap,
                    currentFullCell,
                          SiGlobalVars.Instance.multiTransHeadings,
                         out error_occured, out reached_empty_space_area);

                if (reached_empty_space_area == true)
                    return;


                if (error_occured == true)
                    return;


                //iteration codes===============
                if (nextCell == null)
                    nextCell = firstFullCell.Next;
                else
                    nextCell = nextCell.Next;

                nextFullCell = eXCEL_HELPER.return_full_merg_cell(nextCell);

                i++;
            } while (i <= totalNoUsedColumns);


            multiTransWraps.Add(multiTransWrap);

        }

        internal static bool ReExtract_siteNos_from_multipleTransactions_cells()
        {


            foreach (var multiWrap in SiGlobalVars.Instance.multiTransWraps)
            {
                //site no cell will be after total time worked cell
                //so


                var siteNoCell = multiWrap.totalTimeWorked.fullCell.Next;

                String shortSiteNo = siteNoCell.Value;

                //when totla worktime is available then  it is
                //must to have a valid site no
                //exception is in the case of holidays and friday
                //in holidays and fridays even though there is no valid worktime
                //we might need to mention site no.s
                if (TotalWorkTime_for_employee_is_zeroOrNull(multiWrap.totalTimeWorked)
                        == false)
                {
                    if (shortSiteNo == null)
                    {
                        MessageBox.Show("Cell value of " + siteNoCell.Address
                            + "is null");
                        return false;
                    }

                    shortSiteNo = shortSiteNo.Trim();
                    if (CommonOperations.Given_siteNo_is_validShortSiteNo(shortSiteNo,
                        SiGlobalVars.Instance.DEFAULT_MECHANICAL_SITE_CHAR)
                    == false)
                    {
                        //worktime is non zero or not null
                        //then site no is compulsary
                        MessageBox.Show("This short site No = " + shortSiteNo + " for the employee = "
                       + multiWrap.personnelNo.content + " in Muliple Transaction is not valid");
                        return false;

                    }

                }

                if (CommonOperations.Given_siteNo_is_validShortSiteNo(shortSiteNo, 'M')
                    == true)
                {
                    if (multiWrap.siteNoMechFormat == null)
                        multiWrap.siteNoMechFormat = new MultiTransWrap.SiteNoMechFormat();

                    if (multiWrap.siteNoMechFormat.shortName == null)
                        multiWrap.siteNoMechFormat.shortName = new MultiTransWrap.StrItemWrap();

                    multiWrap.siteNoMechFormat.shortName.content = shortSiteNo;
                    multiWrap.siteNoMechFormat.shortName.fullCell = siteNoCell;
                }



            }
            return true;
        }

        internal static Boolean AutoFill_SiteNos_for_fridaysAndHolidays
            (Form previousForm,  Label label_holidays)
        {

            //first display the Re-extract dialog form
            Boolean autoFillSiteNo = false;
            Form_AutoFillDialog form_AutoFillDialog
                = new Form_AutoFillDialog(previousForm);

            DialogResult dialogResult
                = form_AutoFillDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {

                autoFillSiteNo = form_AutoFillDialog.AutoFillSiteNo;

            }
            else
            {
                MessageBox.Show("AutoFill Site No. Dialog was cancelled." +
                    " Hence this task is interrupted");
                return false;
            }


            if (autoFillSiteNo == true)
            {
                //for fridays and holiday we will fill the site no.s


                if (CommonOperations
                     .Display_holiday_selectorForm(previousForm,
                     label_holidays) == false)
                {
                    MessageBox.Show("Holiday Selector Dialog was cancelled." +
                        " Hence this task is interrupted");
                    return false;
                }

                foreach (var multiWrap in SiGlobalVars.Instance.multiTransWraps)
                {
                    if (shortSiteNo_is_null(multiWrap.siteNoMechFormat)
                        == true)
                    {
                        Boolean holidayOrFriday
                      = MixTimeSheetHandler
                      .Given_date_is_a_holidayOrFriday(multiWrap.date.content.Value,
                      SiGlobalVars.Instance.Holidays);

                        if (holidayOrFriday == false)
                            continue;

                        Boolean ignoreThisAutofill = false;
                        autoFill_SiteNo_for_holidaysAndFridays(multiWrap,
                            out ignoreThisAutofill);

                        if (ignoreThisAutofill == true)
                            continue;

                    }


                }



            }
            return true;
        }

        private static bool shortSiteNo_is_null(MultiTransWrap.SiteNoMechFormat siteNoMechFormat)
        {
            if (siteNoMechFormat == null)
                return true;
            else if (siteNoMechFormat.shortName == null)
                return true;
            else if (siteNoMechFormat.shortName.content == null)
                return true;
            else if (siteNoMechFormat.shortName.fullCell == null)
                return true;
            else if (String.IsNullOrEmpty(siteNoMechFormat.shortName.content)
                         || String.IsNullOrWhiteSpace(siteNoMechFormat.shortName.content))
                return true;

            return false;
        }

        private static void autoFill_SiteNo_for_holidaysAndFridays
            (MultiTransWrap multiWrap, out Boolean ignoreThisAutoFill)
        {
            ignoreThisAutoFill = false;
            if (given_employee_is_present_twoDays_beforeAndAfter(multiWrap) == true)
            {
                //it means this guy didn't go for vacation and all
                //it means due to friday or holiday his attendance was not marked in the biometric


                String proposedNewShortSiteNo 
                    = get_givenEmployee_last_valid_shortSiteNo(multiWrap,out ignoreThisAutoFill);

                if (ignoreThisAutoFill == true)
                    return;

                //now write the new site no
                if (multiWrap.siteNoMechFormat == null)
                    multiWrap.siteNoMechFormat = new MultiTransWrap.SiteNoMechFormat();
                if (multiWrap.siteNoMechFormat.shortName == null)
                    multiWrap.siteNoMechFormat.shortName = new MultiTransWrap.StrItemWrap();

                multiWrap.siteNoMechFormat.shortName.content = proposedNewShortSiteNo;
                multiWrap.siteNoMechFormat.shortName.fullCell
                    = multiWrap.totalTimeWorked.fullCell.Next;
                multiWrap.siteNoMechFormat.shortName.fullCell.Value
                    = proposedNewShortSiteNo;

            }

        }

        private static string get_givenEmployee_last_valid_shortSiteNo
            (MultiTransWrap thisEmpWrap, out Boolean ignoreThisAutofill)
        {
            ignoreThisAutofill = false;
            var completeAttendOfThisEmp = SiGlobalVars.Instance.multiTransWraps
                  .Where(i => i.personnelNo == thisEmpWrap.personnelNo).ToList();
            Boolean stopBefore = false;
            Boolean stopAfter = false;

            for (int i = 1; i <= 30; i++)
            {

                var currentDate = thisEmpWrap.date.content.Value;
                if (stopBefore == false)
                {
                    var beforeDateToCheck = currentDate.AddDays(-i);
                    if (beforeDateToCheck.Month == thisEmpWrap.date.content.Value.Month)
                    {
                        String siteNo =
                            get_a_valid_shortSiteNo_from_thisDate
                            (beforeDateToCheck,thisEmpWrap,completeAttendOfThisEmp);
                        if (siteNo != null)
                            return siteNo;
                    }
                    else
                    {
                        stopBefore = true;
                    }

                }

                if (stopAfter == false)
                {
                    var afterDateToCheck = currentDate.AddDays(i);
                    if (afterDateToCheck.Month == thisEmpWrap.date.content.Value.Month)
                    {
                        String siteNo = 
                            get_a_valid_shortSiteNo_from_thisDate
                            (afterDateToCheck,thisEmpWrap,completeAttendOfThisEmp);
                        if (siteNo != null)
                            return siteNo;
                    }
                    else
                    {
                        stopAfter = true;
                    }

                }

                if (stopBefore == true & stopAfter == true)
                    break;

            }
            ignoreThisAutofill = true;
            return null;
        }

        private static String get_a_valid_shortSiteNo_from_thisDate
            ( DateTime dateToCheck,
            MultiTransWrap thisEmpWrap,
            List<MultiTransWrap> completeAttendOfThisEmp)
        {

            

            foreach (var dateEmp in completeAttendOfThisEmp)
            {
                if (DateTimeHandler
                    .Compare_dates_only(dateToCheck, dateEmp.date.content.Value))
                {
                    if (TotalWorkTime_for_employee_is_zeroOrNull
                 (dateEmp.totalTimeWorked) == true)
                        return null;//if the total time is null 
                                    //then there is no use to get site no in this particular date

                    if (shortSiteNo_is_null(dateEmp.siteNoMechFormat)
                            == true)
                        return null;

                    if (CommonOperations
                          .Given_siteNo_is_validShortSiteNo
                          (dateEmp.siteNoMechFormat.shortName.content,
                          SiGlobalVars.Instance.DEFAULT_MECHANICAL_SITE_CHAR)
                          == true)
                        return dateEmp.siteNoMechFormat.shortName.content;

                }
            }

            return null;
        }

        private static bool given_employee_is_present_twoDays_beforeAndAfter
            (MultiTransWrap givenEmployeeWrap)
        {
            List<DateTime> ThreeDatesBeforeToCheck
                = get_3dates_beforeOrAfter(givenEmployeeWrap.date.content.Value,
                before: true);

            Boolean atleastOncepresent_before3Days
                = check_if_employee_present_atleastOnce_in_givenDates
                (ThreeDatesBeforeToCheck, givenEmployeeWrap);


            List<DateTime> ThreeDatesAfterToCheck
                 = get_3dates_beforeOrAfter(givenEmployeeWrap.date.content.Value, false);

            Boolean atleastOncepresent_after3Days
                = check_if_employee_present_atleastOnce_in_givenDates
                (ThreeDatesAfterToCheck, givenEmployeeWrap);

            if (atleastOncepresent_before3Days == true
                && atleastOncepresent_after3Days == true)
                return true;

            return false;

        }

        private static List<DateTime> get_3dates_beforeOrAfter(DateTime currentDate,
            bool before)
        {
            List<DateTime> dates = new List<DateTime>();
            int valueOperator;

            //means before 3 days
            for (int i = 1; i <= 3; i++)
            {

                if (before == true)
                    valueOperator = -i;//subtract days
                else
                    valueOperator = i;//add days

                var assumedBeforeOrAfterDate = currentDate.AddDays(valueOperator);
                if (assumedBeforeOrAfterDate.Month != currentDate.Month)
                    break;//means we are on the extremes of the month

                dates.Add(assumedBeforeOrAfterDate);

            }

            return dates;

        }

        private static bool check_if_employee_present_atleastOnce_in_givenDates
            (List<DateTime> datesToCheck, MultiTransWrap givenEmployeeWrap)
        {
            if (datesToCheck.Count == 0)
            {
                //means special case
                //sometimes 3 days before cannot be taken
                //for example
                //asumme a date = 01/aug/2019
                //if we take a day before, it will be a day on july month
                //which we can't consider as it is the time sheet of an employee of august month
                //ie if the list count =0
                //means special case
                return true;


            }

            foreach (var date in datesToCheck)
            {

                foreach (var tempMultiWrap in SiGlobalVars.Instance.multiTransWraps)
                {
                    if (tempMultiWrap.personnelNo != givenEmployeeWrap.personnelNo)
                        continue;

                    if (DateTimeHandler.Compare_dates_only(date, tempMultiWrap.date.content.Value)
                        == false)
                        continue;

                    //if both are equal then
                    //check if the total time worked is valid
                    if (TimeSpanHelper
                         .GivenTimeSpan_is_zeroOrNull(tempMultiWrap.totalTimeWorked.content)
                         == false)
                        return true;


                }


            }

            return false;//not at all present
        }

        internal static bool TotalWorkTime_for_employee_is_zeroOrNull
            (MultiTransWrap.TimeSpanItemWrap totalWorkTime)
        {
            if (totalWorkTime == null)
                return true;

            if (totalWorkTime.content == null)
                return true;

            bool result = TimeSpanHelper.GivenTimeSpan_is_zeroOrNull(totalWorkTime.content);
            return result;
        }

        private void feed_time_data_to_dataWrap(ref MultiTransWrap.TimeSpanItemWrap time_data,
     EXCEL_HELPER eXCEL_HELPER, Excel.Range fullCell,
     HeadingWrap heading, DateTime date_of_time)
        {
            //that is employee no
            if (time_data == null)
                time_data = new MultiTransWrap.TimeSpanItemWrap();
            String extractedDate_in_string
                = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
            TimeSpan result_time;

            if (TimeSpan.TryParse(extractedDate_in_string,
                out result_time)
            == true)
                time_data.content = result_time;
            else
                time_data.content = null;

            time_data.fullCell = fullCell;
            time_data.heading = heading;
            time_data.contentInString =
                           eXCEL_HELPER.get_value_of_merge_cell(fullCell);
        }

        private void feed_time_data_to_dataWrap(ref MultiTransWrap.DateItemWrap time_data,
        EXCEL_HELPER eXCEL_HELPER, Excel.Range fullCell,
        HeadingWrap heading, DateTime date_of_time)
        {
            //that is employee no
            if (time_data == null)
                time_data = new MultiTransWrap.DateItemWrap();
            String extractedDate_in_string
                = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
            DateTime result_time;

            if (DateTime.TryParseExact(extractedDate_in_string,
                "HH:mm", CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AdjustToUniversal,
                out result_time)
            == true)
                time_data.content = DateTimeHandler
                    .mix_different_date_and_time(date_of_time, result_time);
            else
                time_data.content = null;

            time_data.fullCell = fullCell;
            time_data.heading = heading;
            time_data.contentInString =
                           eXCEL_HELPER.get_value_of_merge_cell(fullCell);
        }

        private Boolean feed_datas_of_single_row
            (ref MultiTransWrap multiTransWrap,
           Excel.Range fullCell, MultiTransHelper.MultiHeadings headings,
           out Boolean error_occured, out Boolean reached_empty_space_or_invalid_data)
        {
            error_occured = false;
            reached_empty_space_or_invalid_data = false;
            if (fullCell.Column > headings.totalTimeWorked.fullCell.Column)
                return false;

            //Todo: should check there is no merged cells in the timesheet data in future
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);


            foreach (HeadingWrap heading in headings)
            {
                if (heading.Equals(headings.sheetHeading))
                    continue;//because the title of the time sheet that is
                //" Multiple Transaction" which we don't want to iterate
                //for all other headings we need get the corresponding datas
                //from the row below of the heading

                if (fullCell.Column == heading.fullCell.Column)
                {


                    if (eXCEL_HELPER.is_this_a_merged_cell(fullCell)
                         == true)
                    {

                        //we don't enterain merge cells here
                        error_occured = true;
                        MessageBox.Show("Merge cells were found under the heading " +
                            heading.headingName + ". Merged Cell Address = "
                            + fullCell.Address.ToString());

                        return false;

                    }

                    //same column number means the current cell is 
                    //the value for this heading
                    if (heading.Equals(headings.personnelNo))
                    {
                        //that is this particular cell is personal no data
                        String extractedEmployeeNo = eXCEL_HELPER.get_value_of_merge_cell(fullCell);

                        if (CommonOperations.employeeNo_is_valid(extractedEmployeeNo)
                         == false)
                        {
                            MessageBox.Show("Employee No. is empty or invalid in the cell = "
                                + fullCell.Address.ToString() +
                                " Row No = " + fullCell.Row
                                + " Thus Data Extraction is going to stop with this Row");

                            reached_empty_space_or_invalid_data = true;
                            return false;
                        }
                        if (multiTransWrap.personnelNo == null)
                            multiTransWrap.personnelNo = new MultiTransWrap.StrItemWrap();
                        multiTransWrap.personnelNo.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        multiTransWrap.personnelNo.fullCell = fullCell;
                        multiTransWrap.personnelNo.heading = heading;
                        return true;
                    }
                    else if (heading.Equals(headings.firstName))
                    {
                        if (multiTransWrap.firstName == null)
                            multiTransWrap.firstName = new MultiTransWrap.StrItemWrap();
                        String extractedName = eXCEL_HELPER.get_value_of_merge_cell(fullCell);

                        if (CommonOperations.name_is_valid(extractedName)
                            == false)
                        {
                            MessageBox.Show("Name is empty or invalid in the cell = "
                                + fullCell.Address.ToString() +
                                " Row No = " + fullCell.Row
                                + " Thus Data Extraction is going to stop with this Row");

                            reached_empty_space_or_invalid_data = true;
                            return false;
                        }
                        multiTransWrap.firstName.content = extractedName;
                        multiTransWrap.firstName.fullCell = fullCell;
                        multiTransWrap.firstName.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.lastName))
                    {
                        if (multiTransWrap.lastName == null)
                            multiTransWrap.lastName = new MultiTransWrap.StrItemWrap();

                        multiTransWrap.lastName.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        multiTransWrap.lastName.fullCell = fullCell;
                        multiTransWrap.lastName.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.position))
                    {
                        if (multiTransWrap.position == null)
                            multiTransWrap.position = new MultiTransWrap.StrItemWrap();
                        multiTransWrap.position.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        multiTransWrap.position.fullCell = fullCell;
                        multiTransWrap.position.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.department))
                    {
                        if (multiTransWrap.department == null)
                            multiTransWrap.department = new MultiTransWrap.StrItemWrap();
                        multiTransWrap.department.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        multiTransWrap.department.fullCell = fullCell;
                        multiTransWrap.department.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.date))
                    {
                        if (multiTransWrap.date == null)
                            multiTransWrap.date = new MultiTransWrap.DateItemWrap();
                        String extractedDate_in_string = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        DateTime extractedDAte;

                        if (DateTime.TryParse(extractedDate_in_string, out extractedDAte)
                        == true)
                            multiTransWrap.date.content = extractedDAte;
                        else
                        {
                            MessageBox.Show("Found invalid date in cell = " +
                               fullCell.Address);
                            error_occured = true;
                            return false;
                        }
                        multiTransWrap.date.contentInString =
                            eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        multiTransWrap.date.fullCell = fullCell;
                        multiTransWrap.date.heading = heading;

                        //reaching the total over time
                        //as you know after total over time it is overtime datas
                        //so we need to break from this iteration now

                        return true;

                    }
                    else if (heading.Equals(headings.checkInTime1))
                    {

                        feed_time_data_to_dataWrap(ref multiTransWrap.checkInTime1,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);


                        return true;

                    }
                    else if (heading.Equals(headings.checkOutTime1))
                    {
                        feed_time_data_to_dataWrap(ref multiTransWrap.checkOutTime1,
                             eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.workingTime1))
                    {
                        feed_time_data_to_dataWrap(ref multiTransWrap.workingTime1,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.checkInTime2))
                    {
                        feed_time_data_to_dataWrap(ref multiTransWrap.checkInTime2,
                             eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.checkOutTime2))
                    {
                        feed_time_data_to_dataWrap(ref multiTransWrap.checkOutTime2,
                             eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.workingTime2))
                    {
                        feed_time_data_to_dataWrap(ref multiTransWrap.workingTime2,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.totalTimeWorked))
                    {
                        feed_time_data_to_dataWrap(ref multiTransWrap.totalTimeWorked,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }


                }
            }
            return false;
        }



    }
}
