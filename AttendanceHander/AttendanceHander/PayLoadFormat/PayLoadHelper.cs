using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static AttendanceHander.MixTimeSheetHandler;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander.PayLoadFormat
{
    public class PayLoadHelper
    {
        private Excel.Workbook workbook;

        public PayLoadHelper( Excel.Workbook workbook)
        {
            this.workbook = workbook;
        }

        public class PayloadHeadings : IEnumerable<HeadingWrap>
        {

            public HeadingWrap company = new HeadingWrap("Company");
            public HeadingWrap date = new HeadingWrap("Date");
            public HeadingWrap section = new HeadingWrap("Section");
            public HeadingWrap job = new HeadingWrap("Job");
            //========below headings are for iteration==========
            public HeadingWrap serialNo = new HeadingWrap("Sl.No.");
            public HeadingWrap code = new HeadingWrap("Code");
            public HeadingWrap name = new HeadingWrap("Name");
            public HeadingWrap design = new HeadingWrap("Design.");
            public HeadingWrap job_siteNo = new HeadingWrap("Shift/Job");
            public HeadingWrap workTime = new HeadingWrap("Work Time");
            public HeadingWrap noBreak = new HeadingWrap("No Break");
            public HeadingWrap overTime = new HeadingWrap("Over Time");


            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {company,date,section,job, serialNo,
                    code,name,design,job_siteNo,
                    workTime,noBreak,overTime}.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();

            }
        }




        private void extract_data_from_sheets(out Boolean error_found)
        {
            error_found = false;
            int totalDaysInMonth = get_total_days_in_this_month();
            if (totalDaysInMonth < 0)
            {
                error_found = true;
                return;
            }

            loop_through_each_date_sheets_in_order(totalDaysInMonth, out error_found);
            if (error_found == true)
                return;




        }

        private bool loop_through_each_date_sheets_in_order(int totalDaysinMonth,
            out Boolean error_found)
        {
            error_found = false;
            //first find the first sheet.



            Excel.Worksheet firstSheet = null;

            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {

                if (sheet.Name.Trim() == "1")
                {
                    firstSheet = sheet;
                    break;
                }


            }

            if (firstSheet == null)
            {
                MessageBox.Show("Couldn't find sheet no = 1 in PayloadFormat");
                error_found = true;
                return false;
            }


            for (int day = 1; day <= totalDaysinMonth; day++)
            {
                Excel.Worksheet currentSheet;
                //sheets should be in order

                if (day != 1)
                {
                    currentSheet = firstSheet.Next();
                    if (currentSheet.Name.Trim() != day.ToString())
                    {
                        MessageBox.Show("Sheets might not be in order; Expected sheet = "
                            + day.ToString() + "; but the sheet obtained = " + currentSheet.Name);
                        error_found = true;
                        return false;
                    }
                }
                else
                {
                    currentSheet = firstSheet;
                }

                PayLoadWrap.Day payLoadWrapDay = new PayLoadWrap.Day( ref currentSheet);


                find_headings(ref SiGlobalVars.
                 Instance.payLoadHeadings, out error_found, ref currentSheet);


                if (error_found == true)
                    return false;
                //now we got all the headings
                //connect heading and datas together

                //now that we got all headings
                //we need to start with the rows


                
                //first Pre-table datas are headings like company, date, section, job


                read_preTable_datas(out error_found, ref payLoadWrapDay);

                if (error_found == true)
                {
                    return false;
                }

                //Once we got pre-table datas like company, date, section, job etc
                //we need to find the datas for each employee.
                read_each_rows_of_data(out error_found, ref payLoadWrapDay, ref currentSheet);
                if (error_found == true)
                    return false;

                //finaly add to global variable
                SiGlobalVars.Instance.payLoadWrap.days.Add(payLoadWrapDay);

            }



            return true;
        }



        private void read_preTable_datas(out bool error_found, ref PayLoadWrap.Day payLoadWrapDay)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(payLoadWrapDay.sheet);
            //===for company
            Excel.Range company_data = eXCEL_HELPER.return_next_adjacent_range(SiGlobalVars.Instance
                .payLoadHeadings.company.fullCell);

            if (payLoadWrapDay.company == null)
                payLoadWrapDay.company = new PayLoadWrap.StrItemWrap();
            payLoadWrapDay.company.fullCell = company_data;
            payLoadWrapDay.company.content = eXCEL_HELPER.get_value_of_merge_cell(company_data);

            //===for date
            Excel.Range date_data = eXCEL_HELPER.return_next_adjacent_range(SiGlobalVars.Instance
                .payLoadHeadings.date.fullCell);
            if (payLoadWrapDay.date == null)
                payLoadWrapDay.date = new PayLoadWrap.DateItemWrap();
            payLoadWrapDay.date.fullCell = date_data;
            payLoadWrapDay.date.contentInString = eXCEL_HELPER.get_value_of_merge_cell(date_data);

            //===for section
            Excel.Range section_data = eXCEL_HELPER.return_next_adjacent_range(SiGlobalVars.Instance
               .payLoadHeadings.section.fullCell);
            if (payLoadWrapDay.section == null)
                payLoadWrapDay.section = new PayLoadWrap.StrItemWrap();
            payLoadWrapDay.section.fullCell = section_data;
            payLoadWrapDay.section.content = eXCEL_HELPER.get_value_of_merge_cell(section_data);

            //===for job
            Excel.Range job_data = eXCEL_HELPER.return_next_adjacent_range(SiGlobalVars.Instance
               .payLoadHeadings.job.fullCell);
            if (payLoadWrapDay.job == null)
                payLoadWrapDay.job = new PayLoadWrap.StrItemWrap();
            payLoadWrapDay.job.fullCell = job_data;
            payLoadWrapDay.job.content = eXCEL_HELPER.get_value_of_merge_cell(job_data);

        }

        private void read_each_rows_of_data(out bool error_occured,
           ref PayLoadWrap.Day payloadWrapDay,
           ref Excel.Worksheet currentSheet)
        {
            error_occured = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(currentSheet);

            //now we have to read the rows
            //our beginning position to start reading the rows will be
            //below the personal Number heading cell

            Excel.Range serialNoHeading = SiGlobalVars.Instance
                .payLoadHeadings.serialNo.fullCell;

            int lastHeadingRow = EXCEL_HELPER.get_last_row_no_of_a_merged_cell(serialNoHeading);
            //now go to below adjacent cell to personnal no.
            int firstDataRowNo = lastHeadingRow + 1;



            foreach (Excel.Range row in currentSheet.UsedRange.Rows)
            {
                int currentRow = row.Row;
                //our first data row starts from firstDataRowCell
                //so skip the rows above (which are headings)
                if (currentRow < firstDataRowNo)
                    continue;

                //read row will fail at the end of time sheet
                //we need to stop the iteration after plumber no 50
                //after that it is just empty space
                //so if reached_empty_space_area = true, we came accross the empty space
                //thus this iteration can be stopped.

                Boolean reached_empty_space_area = false;
                read_row(row,
                    ref payloadWrapDay,
                    out error_occured, out reached_empty_space_area,ref currentSheet);
                if (error_occured == true)
                    return;
                if (reached_empty_space_area == true)
                    break;

            }

        }

        private void read_row(Excel.Range row,
            ref PayLoadWrap.Day payLoadWrapDay,
            out bool error_occured, out bool reached_empty_space_area,
            ref Excel.Worksheet currentSheet)
        {
            reached_empty_space_area = false;
            error_occured = false;
            //one way to identify whether we are in empty space
            //that means whether we already passed 50 numbers of plumbers
            //is by detecting if serial no,employee no and name etc are empty
            //if the serial no, employee no and name is empty means 
            //we have reached the end of the time sheet
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(currentSheet);
            int rowIndex = row.Row;

            Excel.Range firstCell = currentSheet.Cells[rowIndex, 1];
            Excel.Range firstFullCell = eXCEL_HELPER.return_full_merg_cell(firstCell);
            Excel.Range nextCell = null;

            Excel.Range nextFullCell = firstFullCell;
            int totalNoUsedColumns = currentSheet.UsedRange.Columns.Count;

            int i = 1;
            //to check if we have reached the empty space or blank area after 
            // if no name or employee no is found in this row
            // then we can say we reached the empty space

            PayLoadWrap.Day.Employee payLoadWrapDayEmpl = new PayLoadWrap.Day.Employee();
            do
            {
                //first nextFullCell is firstFullCell
                //so 
                var currentFullCell = nextFullCell;




                Boolean result1;
                result1 = feed_datas_of_single_row(ref payLoadWrapDayEmpl,
                    currentFullCell,
                          SiGlobalVars.Instance.payLoadHeadings,
                         out error_occured, out reached_empty_space_area,
                         ref currentSheet);

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


            payLoadWrapDay.employees.Add(payLoadWrapDayEmpl);
        }

        private bool feed_datas_of_single_row(ref PayLoadWrap.Day.Employee payLoadWrapDayEmpl,
            Excel.Range fullCell, PayloadHeadings payLoadHeadings,
            out bool error_occured, out bool reached_empty_space_or_invalid_data,
           ref Excel.Worksheet currentSheet)
        {
            error_occured = false;
            reached_empty_space_or_invalid_data = false;
            if (fullCell.Column > payLoadHeadings.overTime.fullCell.Column)
                return false;


            //Todo: should check there is no merged cells in the timesheet data in future
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(currentSheet);


            foreach (HeadingWrap heading in payLoadHeadings)
            {


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
                    if (heading.Equals(payLoadHeadings.serialNo))
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
                        if (payLoadWrapDayEmpl.serialNo == null)
                            payLoadWrapDayEmpl.serialNo = new PayLoadWrap.StrItemWrap();
                        payLoadWrapDayEmpl.serialNo.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.serialNo.fullCell = fullCell;
                        payLoadWrapDayEmpl.serialNo.heading = heading;
                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.code))
                    {
                        if (payLoadWrapDayEmpl.code == null)
                            payLoadWrapDayEmpl.code = new PayLoadWrap.StrItemWrap();
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
                        payLoadWrapDayEmpl.code.content = extractedName;
                        payLoadWrapDayEmpl.code.fullCell = fullCell;
                        payLoadWrapDayEmpl.code.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.name))
                    {
                        if (payLoadWrapDayEmpl.name == null)
                            payLoadWrapDayEmpl.name = new PayLoadWrap.StrItemWrap();

                        payLoadWrapDayEmpl.name.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.name.fullCell = fullCell;
                        payLoadWrapDayEmpl.name.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.design))
                    {
                        if (payLoadWrapDayEmpl.design == null)
                            payLoadWrapDayEmpl.design = new PayLoadWrap.StrItemWrap();
                        payLoadWrapDayEmpl.design.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.design.fullCell = fullCell;
                        payLoadWrapDayEmpl.design.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.job_siteNo))
                    {
                        if (payLoadWrapDayEmpl.job_siteNo == null)
                            payLoadWrapDayEmpl.job_siteNo = new PayLoadWrap.StrItemWrap();
                        payLoadWrapDayEmpl.job_siteNo.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.job_siteNo.fullCell = fullCell;
                        payLoadWrapDayEmpl.job_siteNo.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.workTime))
                    {
                        if (payLoadWrapDayEmpl.workTime == null)
                            payLoadWrapDayEmpl.workTime = new PayLoadWrap.DecimalItemWrap();
                        payLoadWrapDayEmpl.workTime.contentInStr
                            = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.workTime.fullCell = fullCell;
                        payLoadWrapDayEmpl.workTime.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.noBreak))
                    {
                        if (payLoadWrapDayEmpl.workTime == null)
                            payLoadWrapDayEmpl.workTime = new PayLoadWrap.DecimalItemWrap();
                        payLoadWrapDayEmpl.workTime.contentInStr
                            = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.workTime.fullCell = fullCell;
                        payLoadWrapDayEmpl.workTime.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(payLoadHeadings.overTime))
                    {
                        if (payLoadWrapDayEmpl.overTime == null)
                            payLoadWrapDayEmpl.overTime = new PayLoadWrap.DecimalItemWrap();
                        payLoadWrapDayEmpl.overTime.contentInStr
                            = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        payLoadWrapDayEmpl.overTime.fullCell = fullCell;
                        payLoadWrapDayEmpl.overTime.heading = heading;
                        return true;
                    }

                }
            }
            return false;
        }


        private int get_total_days_in_this_month()
        {
            //no of sheets depends upon number of days in a month
            //eg febraury = 28 
            //so first we need to understand which month and year
            //we can refer it from the multitranswrap
            //assume any date from and employee and thus we can get the date and month
            DateTime assumedDate = SiGlobalVars.Instance
                .multiTransWraps.First().date.content.Value;

            if (assumedDate == null)
            {
                MessageBox.Show("Date in Multiple Transaction is null; Cell Address = "
                    + SiGlobalVars.Instance
                .multiTransWraps.First().date.fullCell.Address +
                "; Content = ");
                return -1;
            }
            int noOfDays = DateTime.DaysInMonth(assumedDate.Year, assumedDate.Month);


            return noOfDays;
        }



        public void MAIN_understand_the_excel_sheet(out Boolean error_found)
        {
            error_found = false;
            if (SiGlobalVars.Instance.payLoadHeadings == null)
            {
                SiGlobalVars.Instance.payLoadHeadings
                    = new PayLoadHelper.PayloadHeadings();
            }

            if (SiGlobalVars.Instance.payLoadWrap == null)
                SiGlobalVars.Instance.payLoadWrap
                    = new PayLoadWrap();


            extract_data_from_sheets(out error_found);





        }



        internal static MixTimeSheetHandler.WorkTimeCalculatedWarp Calculate_worktime_from_bioTotalWorkTime
            (TimeSpanItemWrap totalTimeWorked, bool thisDate_is_fridayOrHoliday)
        {
            MixTimeSheetHandler.WorkTimeCalculatedWarp workTimeCalculated
                = new MixTimeSheetHandler.WorkTimeCalculatedWarp();

            //TODO: Worktime is hardcorded that is 8 hours
            //in future we need it to flexible.


            if (totalTimeWorked.content.Value.Hours < SiGlobalVars.Instance.DEFAULT_WORKING_HOURS)
            {
                //i.e case when work time is below 8 hours ...like 7 hours
                //in this case we only have worktime and the overtime is zero.

                if (workTimeCalculated.workTimeHours == null)
                    workTimeCalculated.workTimeHours = new WorkTimeCalculatedWarp.Wrap();
                if (workTimeCalculated.overTime == null)
                    workTimeCalculated.overTime = new WorkTimeCalculatedWarp.Wrap();
                decimal resultTotalHours;

                roundup_hours_based_on_minutes(totalTimeWorked.content.Value,
                    out resultTotalHours);

                if (resultTotalHours == -1)
                {
                    MessageBox.Show("Worktime Hours was found to be -1");
                    return null;
                }
                if (thisDate_is_fridayOrHoliday == true)
                {
                    //ie friday or holiday means
                    //what ever duration he worked will go to overtime instead of normal work time
                    workTimeCalculated.workTimeHours.content = 0;
                    workTimeCalculated.overTime.content = resultTotalHours;

                }
                else
                {
                    //ie not friday or holiday
                    //means the time worked will go into normal work time
                    workTimeCalculated.workTimeHours.content = resultTotalHours;
                    workTimeCalculated.overTime.content = 0;
                }



            }
            else if (totalTimeWorked.content.Value.Hours > SiGlobalVars.Instance.DEFAULT_WORKING_HOURS)
            {
                //ie case when total biometric time is greater than 8 hours ...like (13 hours in biometric)
                if (workTimeCalculated.workTimeHours == null)
                    workTimeCalculated.workTimeHours = new WorkTimeCalculatedWarp.Wrap();
                if (workTimeCalculated.overTime == null)
                    workTimeCalculated.overTime = new WorkTimeCalculatedWarp.Wrap();

                decimal resultTotalHours;

                roundup_hours_based_on_minutes(totalTimeWorked.content.Value,
                    out resultTotalHours);

                if (workTimeCalculated.workTimeHours.content == -1)
                {
                    MessageBox.Show("Worktime Hours was found to be -1");
                    return null;
                }

                if (thisDate_is_fridayOrHoliday == true)
                {
                    //ie friday or holiday means
                    //what ever duration he worked will go to overtime instead of normal work time
                    workTimeCalculated.workTimeHours.content = 0;
                    workTimeCalculated.overTime.content = resultTotalHours;

                }
                else
                {
                    //ie not friday or holiday
                    //means the time worked will go into normal work time
                    workTimeCalculated.workTimeHours.content = SiGlobalVars.Instance.DEFAULT_WORKING_HOURS;
                    workTimeCalculated.overTime.content
                        = (resultTotalHours -
                        (SiGlobalVars.Instance.DEFAULT_WORKING_HOURS +
                        SiGlobalVars.Instance.DEFAULT_BREAK_HOURS));
                }



            }
            return workTimeCalculated;
        }

        private static Boolean roundup_hours_based_on_minutes(TimeSpan timeSpan,
            out decimal resultHours)
        {
            resultHours = -1;

            if (timeSpan == null)
                return false;

            if (timeSpan.Minutes >= 0 && timeSpan.Minutes <= 15)
            {
                resultHours = timeSpan.Hours;
                return true;


            }
            else if (timeSpan.Minutes >= 16 && timeSpan.Minutes <= 45)
            {
                resultHours = timeSpan.Hours + 0.5M;
                return true;


            }
            else if (timeSpan.Minutes >= 46 && timeSpan.Minutes <= 60)
            {
                resultHours = timeSpan.Hours + 1;
                return true;
            }

            return false;

        }
     
        private Boolean find_headings(ref PayloadHeadings payLoadHeadings,
            out bool error_found,
           ref Excel.Worksheet currentSheet)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(currentSheet);
            foreach (HeadingWrap heading in payLoadHeadings)
            {

                if (heading.fullCell != null)
                    continue;


                List<Excel.Range> search_results = new List<Excel.Range>();
                search_results =
                    eXCEL_HELPER.find_fix_column_heading(heading.headingName.Trim(),
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

                    Excel.Range fullcell = eXCEL_HELPER
                        .return_full_merg_cell(search_results[0]);
                    heading.fullCell = fullcell;
                }
                else
                {
                    //TODO: if more than one search results
                    //then show error and abort
                    //in future we can implement some codes to 
                    //filter out the search results.

                    MessageBox.Show("More than one Search results " +
                        "was found for the heading; Heading = " + heading.headingName
                        + "; Cell Address = ");
                    error_found = true;
                    return false;

                }

            }



            return true;

        }
    }
}
