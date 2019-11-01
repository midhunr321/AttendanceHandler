using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    public class MepStyleHelper
    {

        private Excel.Worksheet worksheet;
        private Excel.Workbook workbook;

        public MepStyleHelper(Excel.Workbook workbook, Excel.Worksheet worksheet)
        {
            this.workbook = workbook;
            this.worksheet = worksheet;
        }

        public class Headings : IEnumerable<HeadingWrap>
        {
            public HeadingWrap sheetHeading =
                new HeadingWrap("Plumbers - Time Sheet");
            public HeadingWrap serialNo = new HeadingWrap("S. No");
            public HeadingWrap code = new HeadingWrap("Code");
            public HeadingWrap name = new HeadingWrap("Name");
            public HeadingWrap designation = new HeadingWrap("Design");
            public HeadingWrap siteNO = new HeadingWrap("Site Nos.");
            public HeadingWrap totalOvertime = new HeadingWrap("Total Over Time");
            public HeadingWrap date = new HeadingWrap("Date:");
            public Dictionary<int, HeadingWrap> overtimeDays;

            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {sheetHeading,serialNo,
                    code,name,designation,siteNO,
                    totalOvertime,date }.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
        public void MAIN_understand_the_excel_sheet(out Boolean error_found)
        {
             error_found = false;

            if (SiGlobalVars.Instance.mepStyleHeadings == null)
            {
                SiGlobalVars.Instance.mepStyleHeadings = new Headings();
            }

            find_headings_except_overtimeDates(ref SiGlobalVars.
                Instance.mepStyleHeadings,out error_found);

            if (error_found == true)
                return;
            //if all the headings
            //are found, it means the opened excel file is Mep style 
            //plumbers time sheet
            understand_the_month_and_year_of_the_sheet();
            if (find_overtime_dates_headings()
                 == false)
                return;


            //now we got all the headings
            //connect heading and datas together

            //now that we got all headings
            //we need to start with the rows
            Boolean error_ocurred = false;
            read_each_rows_of_data(out error_ocurred);
            if (error_ocurred == true)
                return;

            
        }

        private Boolean plumber_is_on_vacation_or_has_merged_cells
            (Excel.Range fullCell)
        {
            //one of the reasons why a cell would be merged is 
            //becasue the plumber would have gone for vacation
            //and hence the overtime cells would be merged together
            //labelled as "Vacation"k


            if (fullCell.MergeCells == true)
                return true;

            return false;
        }
        private Boolean feed_site_transfer_data_of_a_cell(ref List<DateOvertime> dateOvertime,
           Excel.Range fullCell, MepStyleHelper.Headings headings,
          out Boolean stopThisRowIteration)
        {
            stopThisRowIteration = false;
            //inorder to feed site transfer data
            //first we have to make sure that 
            //the column of the current cell is after the day 31 or day 30  of this month

            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            int thisCellLastColNo =
            eXCEL_HELPER.get_last_column_no_of_a_merge_cell(fullCell);
            int lastOvertimeDayColumn = headings.overtimeDays.Last().Value.fullCell.Column;
            if (thisCellLastColNo > lastOvertimeDayColumn)
            {
                //ie this full cell is after the overtime dates
                //that means now on we need to check for site transfer no details
                //if any cell doesn't contain site transfer details then
                // we need to stop the iteration along this row

                MepStyleSiteNoCodeAnalyzer codeAnalyzer
                        = new MepStyleSiteNoCodeAnalyzer
                        (SiGlobalVars.Instance.mepStyleTimesheetMonthYear);

                MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap extractedDataWrap
                    = new MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap();

                String transferCode = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                extractedDataWrap = codeAnalyzer.analyze_string(transferCode);
                if (extractedDataWrap == null)
                {
                    //if no valid site tranfer code is found, we can break from
                    //iteration of the row.
                    stopThisRowIteration = true;
                    return false;
                }
                else
                {
                    insert_transfer_dates_into_datawrap(ref dateOvertime,
                        extractedDataWrap);
                    return true;
                }


            }
            return false;
        }

        private List<DateOvertime> filter_overtime_for_these_dates(DateTime dateForFilter,
            List<DateOvertime> dateOvertime)
        {
            List<DateOvertime> filtered = new List<DateOvertime>();
            foreach (var item in dateOvertime)
            {
                if (item.date.Equals(dateForFilter))
                    filtered.Add(item);
            }
            return filtered;
        }
        private Boolean insert_transfer_dates_into_datawrap(ref List<DateOvertime> dateOvertime,
            MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap extractedDataWrap)
        {
            DateTime startDate = extractedDataWrap.transferStartDate;
            DateTime endDate = extractedDataWrap.transferEndDate;

            for (var i = startDate; i <= endDate; i = i.AddDays(1))
            {

                //var filteredOvertime
                //    = filter_overtime_for_these_dates(i, dateOvertime);
                //if (filteredOvertime.Count > 1)
                //{
                //    MessageBox.Show("Identical Overtime dates detected in the Overtime Date headings"+
                //        "Issue Overtime Heading Date = " + dateOvertime.Last().heading.headingName);
                //    return false;
                //}

                foreach (var item in dateOvertime)
                {
                    if (item.date.Equals(i.Date))
                    {
                        item.siteNo = extractedDataWrap.siteNo;

                    }
                }

            }

            return true;

        }

        private Boolean feed_overtime_datas_of_a_cell(ref MepStyleWrap mepStyleWrap,
           Excel.Range fullCell, MepStyleHelper.Headings headings,
           out Boolean error_occured)
        {
            error_occured = false;

            if (fullCell.Column > headings.overtimeDays.Last().Value.fullCell.Column)
                return false;//because we limit this iteration before 
                             //till last 30 or 31 days (depending on corresponding months)
                             //and we don't want the iteration after that

            //also overtime date comes only after total overtime heading
            //so before that we can skip

            if (fullCell.Column < headings.totalOvertime.fullCell.Column)
                return false;

            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            foreach (var heading in headings.overtimeDays)
            {

                if (fullCell.Column == heading.Value.fullCell.Column)

                {
                    //same column number means the current cell is 
                    //the data for this heading

                    //if more than one merge cell is found
                    //simply ignore it as they must be in vacation.
                    if (plumber_is_on_vacation_or_has_merged_cells(fullCell)
                 == true)
                        return true;

                    //if no merge cells then
                    var currMonthYear = SiGlobalVars.Instance
                        .mepStyleTimesheetMonthYear;
                    int totalMonthDays = DateTime.DaysInMonth(currMonthYear.Year,
                        currMonthYear.Month);

                    int overtimeDay;
                    if (int.TryParse(heading.Value.headingName, out overtimeDay)
                           == true)
                    {
                        if (mepStyleWrap.dateOvertimes == null)
                            mepStyleWrap.dateOvertimes = new List<DateOvertime>();
                        DateOvertime newItem = new DateOvertime();

                        newItem.overtime
                       = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        newItem.heading = heading.Value;
                        newItem.date_day = heading.Value.headingName;
                        DateTime date = new DateTime(currMonthYear.Year,
                        currMonthYear.Month, overtimeDay);
                        newItem.date = date;
                        mepStyleWrap.dateOvertimes.Add(newItem);

                    }
                    else
                    {
                        MessageBox.Show("Couldn't convert heading name to index." +
                            "Heading name might not be a number");
                        error_occured = true;
                        return false;
                    }


                }
            }
            error_occured = false;
            return true;
        }

       
        private Boolean feed_non_overtime_datas_of_single_row
            (ref MepStyleWrap mepStyleWrap,
           Excel.Range fullCell, MepStyleHelper.Headings headings,
           out Boolean error_occured, out Boolean reached_empty_space_or_invalid_data)
        {
            error_occured = false;
            reached_empty_space_or_invalid_data = false;
            if (fullCell.Column > headings.totalOvertime.fullCell.Column)
                return false;//because we limit this iteration before 
                             //overtime datas and we don't want the iteration to 
                             //run into overtime data columns

            //Todo: should check there is no merged cells in the timesheet data in future
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);


            foreach (HeadingWrap heading in headings)
            {
                if (heading.Equals(headings.sheetHeading))
                    continue;//because the title of the time sheet that is
                //" Plumbers timesheet september 2019"
                //is not required for data extraction

                if (fullCell.Column == heading.fullCell.Column)
                {
                    if (eXCEL_HELPER.is_this_a_merged_cell(fullCell)
                         == true)
                    {
                        //if data under non overtime headings such as serial no
                        //name etc are having merged cells then
                        //some error is there
                        //we don't enterain merge cells here
                        error_occured = true;
                        MessageBox.Show("Merge cells were found under the heading " +
                            heading.headingName +". Merged Cell Address = " 
                            + fullCell.Address.ToString());
                        return false;

                    }

                    //same column number means the current cell is 
                    //the value for this heading
                    if (heading.Equals(headings.serialNo))
                    {
                        //that is this particular cell is serial no data
                        if (mepStyleWrap.serialNo == null)
                            mepStyleWrap.serialNo = new StrItemWrap();
                        mepStyleWrap.serialNo.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        mepStyleWrap.serialNo.fullCell = fullCell;
                        mepStyleWrap.serialNo.heading = heading;
                        return true;
                    }
                    else if (heading.Equals(headings.code))
                    {
                        //that is employee no
                        if (mepStyleWrap.code == null)
                            mepStyleWrap.code = new StrItemWrap();
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
                        mepStyleWrap.code.content = extractedEmployeeNo;
                        mepStyleWrap.code.fullCell = fullCell;
                        mepStyleWrap.code.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.name))
                    {
                        //that is employee no
                        if (mepStyleWrap.name == null)
                            mepStyleWrap.name = new StrItemWrap();
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
                        mepStyleWrap.name.content = extractedName;
                        mepStyleWrap.name.fullCell = fullCell;
                        mepStyleWrap.name.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.designation))
                    {
                        //that is employee no
                        if (mepStyleWrap.designation == null)
                            mepStyleWrap.designation = new StrItemWrap();
                        mepStyleWrap.designation.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        mepStyleWrap.designation.fullCell = fullCell;
                        mepStyleWrap.designation.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.siteNO))
                    {
                        //that is employee no
                        if (mepStyleWrap.siteNo == null)
                            mepStyleWrap.siteNo = new StrItemWrap();
                        mepStyleWrap.siteNo.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        mepStyleWrap.siteNo.fullCell = fullCell;
                        mepStyleWrap.siteNo.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.totalOvertime))
                    {
                        //that is employee no
                        if (mepStyleWrap.totalOvertime == null)
                            mepStyleWrap.totalOvertime = new StrItemWrap();
                        mepStyleWrap.totalOvertime.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        mepStyleWrap.totalOvertime.fullCell = fullCell;
                        mepStyleWrap.totalOvertime.heading = heading;

                        //reaching the total over time
                        //as you know after total over time it is overtime datas
                        //so we need to break from this iteration now

                        return true;

                    }


                }
            }
            return false;
        }

        private bool employeeNo_is_valid(string extractedEmployeeNo)
        {
            if (extractedEmployeeNo == null)
                return false;
            if (String.IsNullOrEmpty(extractedEmployeeNo))
                return false;
            if (String.IsNullOrWhiteSpace(extractedEmployeeNo))
                return false;

            String trimmedEmployeeno = extractedEmployeeNo.Trim();
            if (trimmedEmployeeno.Length < 3)
                return false;

            return true;
        }

        private void read_row(Excel.Range row,
            ref List<MepStyleWrap> mepStyleWraps,
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
            ///atleast_one_result_was_true_in_this_row is used
            //to check if we have reached the empty space or blank area after 
            // if no name or employee no is found in this row
            // then we can say we reached the empty space

            MepStyleWrap mepStyleWrap = new MepStyleWrap();
            do
            {
                //first nextFullCell is firstFullCell
                //so 
                var currentFullCell = nextFullCell;

                Boolean result1;
                result1 = feed_non_overtime_datas_of_single_row(ref mepStyleWrap,
                    currentFullCell,
                          SiGlobalVars.Instance.mepStyleHeadings,
                         out error_occured, out reached_empty_space_area);

                if (reached_empty_space_area == true)
                    return;
                

                if (error_occured == true)
                    return;

                Boolean result2;

                result2 = feed_overtime_datas_of_a_cell(ref mepStyleWrap, currentFullCell,
                       SiGlobalVars.Instance.mepStyleHeadings,
                       out error_occured);
                if (error_occured == true)
                    return;

                //now get the site transfer start and end dates
                Boolean stop_this_row_iteration = false;

                Boolean result3;
                result3 = feed_site_transfer_data_of_a_cell
                      (ref mepStyleWrap.dateOvertimes, currentFullCell,
                         SiGlobalVars.Instance.mepStyleHeadings,
                        out stop_this_row_iteration);

                if (stop_this_row_iteration == true)
                    break;


                if (nextCell == null)
                    nextCell = firstFullCell.Next;
                else
                    nextCell = nextCell.Next;
                nextFullCell = eXCEL_HELPER.return_full_merg_cell(nextCell);

                i++;
            } while (i <= totalNoUsedColumns);


           
                

           
                mepStyleWraps.Add(mepStyleWrap);
            

            //foreach (Excel.Range cell in row.Columns)
            //{
            //    int lastColumnNo = SiGlobalVars.Instance
            //          .mepStyleHeadings.overtimeDays.Last()
            //          .Value.fullCell.Column;

            //    feed_non_overtime_datas_of_single_row(ref mepStyleWraps, cell,
            //        SiGlobalVars.Instance.mepStyleHeadings);
            //}
            //return true;
        }

    
        private void read_each_rows_of_data(out Boolean error_occured)
        {
            error_occured = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            //now we have to read the rows
            //our beginning position to start reading the rows will be
            //below the serial no heading cell

            Excel.Range serialNo = SiGlobalVars.Instance
                .mepStyleHeadings.serialNo.fullCell;


            //now go to below adjacent cell to serial no.
            Excel.Range firstDataRowCell =
                eXCEL_HELPER
                .return_immediate_below_cell(serialNo);

            foreach (Excel.Range row in worksheet.UsedRange.Rows)
            {
                //our first data row starts from firstDataRowCell
                //so skip the rows above (which are headings)
                if (row.Row < firstDataRowCell.Row)
                    continue;

                //read row will fail at the end of time sheet
                //we need to stop the iteration after plumber no 50
                //after that it is just empty space
                //so if read row is false that means we came accross the empty space
                //thus this iteration can be stopped.
                Boolean reached_empty_space_area = false;
                read_row(row,
                    ref SiGlobalVars.Instance.mepStyleWraps,
                    out error_occured, out reached_empty_space_area);
                if (error_occured == true)
                    return;
                if (reached_empty_space_area == true)
                    break;

            }

        }

        private void understand_the_month_and_year_of_the_sheet()
        {
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            var headings = SiGlobalVars.Instance.mepStyleHeadings;
            var dateHeading = headings.date.fullCell;
            Excel.Range date = eXCEL_HELPER.
                return_next_adjacent_range(dateHeading);
            Excel.Range fullcell = eXCEL_HELPER.return_full_merg_cell(date);
            String timesheetDate =
                eXCEL_HELPER.get_value_of_merge_cell(fullcell);

            if (SiGlobalVars.Instance.mepStyleWraps == null)
                SiGlobalVars.Instance.mepStyleWraps = new List<MepStyleWrap>();


            if (SiGlobalVars.Instance.mepStyleTimesheetMonthYear == null)
            {
                SiGlobalVars.Instance.
                    mepStyleTimesheetMonthYear = new DateTime();
            }

            SiGlobalVars.Instance.
                mepStyleTimesheetMonthYear =
                DateTime.Parse(timesheetDate);

        }

        private Boolean overtime_dates_are_in_order_and_valid()
        {

            int no_of_days = DateTime.DaysInMonth(SiGlobalVars.Instance.mepStyleTimesheetMonthYear.Year,
                SiGlobalVars.Instance.mepStyleTimesheetMonthYear.Month);
            int i = 1;
            foreach (var item in SiGlobalVars.Instance.mepStyleHeadings.overtimeDays)
            {
                String expected = i.ToString();
                if (item.Value.headingName != expected)
                {
                    MessageBox.Show("Expected Overtime Heading Date was = " +
                        i.ToString() + " but the actual date heading = " + item.Value.headingName);
                    return false;
                }


                i++;
            }
            return true;

        }

        private Boolean find_overtime_dates_headings()
        {
            var timesheetDate = SiGlobalVars.Instance.
                mepStyleTimesheetMonthYear;
            int currentMonthDaysCount
                 = DateTime.DaysInMonth(timesheetDate.Year,
                 timesheetDate.Month);
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            for (int day = 1; day <= currentMonthDaysCount; day++)
            {
                HeadingWrap newDay = new HeadingWrap(day.ToString());

                if (SiGlobalVars.Instance.mepStyleHeadings.overtimeDays
                    == null)
                    SiGlobalVars.Instance.mepStyleHeadings.overtimeDays
                        = new Dictionary<int, HeadingWrap>();

                SiGlobalVars.Instance.mepStyleHeadings.overtimeDays
                    .Add(day, newDay);
            }

            if (overtime_dates_are_in_order_and_valid()
                   == false)
                return false;

            //to find the over time day cells
            //we need to find the cell range next or adjacent to total overtime

            Excel.Range totalOvertimeHeading = SiGlobalVars.Instance.
                mepStyleHeadings.totalOvertime.fullCell;
            //after the total over time heading 
            //the day 1 is starting.
            //so get the next cell

            Excel.Range day1 = eXCEL_HELPER.
                return_next_adjacent_range(totalOvertimeHeading);
            //now that we got the day 1 overtime heading cell
            //lets put it in the wrap
            SiGlobalVars.Instance.
                mepStyleHeadings.overtimeDays[1].fullCell = day1;

            //now for the other days
            Excel.Range lastDay = day1;
            for (int i = 2; i <= currentMonthDaysCount; i++)
            {
                Excel.Range nextday = eXCEL_HELPER
                    .return_next_adjacent_range(lastDay);
                SiGlobalVars.Instance.mepStyleHeadings
                    .overtimeDays[i].fullCell = nextday;
                if (check_overtime_heading_content_is_valid(nextday, i)
                    == false)
                    return false;
                lastDay = nextday;
            }

            return true;

        }

        private Boolean check_overtime_heading_content_is_valid(Excel.Range cell, int day)
        {
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            String cellContentStr = eXCEL_HELPER.get_value_of_merge_cell(cell);

            StringHandler stringHandler = new StringHandler();
            if (stringHandler
                .is_this_string_alpha_numeric_or_numeric_or_alpha_only(cellContentStr)
                != All_const.str_type.Numeric)
            {
                MessageBox.Show("The Over time heading for " + day + " is not valid");
                return false;
            }
            int cellContentInt;
            if (int.TryParse(cellContentStr, out cellContentInt)
                 == false)
            {
                MessageBox.Show("The Over time heading for " + day + " is not valid");
                return false;
            }

            return true;

        }

        private Boolean find_headings_except_overtimeDates(ref MepStyleHelper.Headings headings,
            out Boolean error_found)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            foreach (HeadingWrap heading in headings)
            {
                List<Excel.Range> temp_heading = new List<Excel.Range>();
                temp_heading =
                    eXCEL_HELPER.find_fix_column_heading(heading.headingName,
                    Excel.XlSearchDirection.xlNext,
                    Excel.XlSearchOrder.xlByRows, false);

                if (temp_heading == null)
                {
                    MessageBox.Show("Couldn't find the heading = "
                        + heading.headingName);
                    return false;
                }
                //TODO: Should carryout the search from the top

                // if the search count is not more than 1 then,
                if (temp_heading != null && temp_heading.Count == 1)
                {
                    Excel.Range fullcell = eXCEL_HELPER
                        .return_full_merg_cell(temp_heading[0]);
                    heading.fullCell = fullcell;
                }
                else
                {
                    //TODO: if more than one search results
                    //we need to filter it out
                    //like check if the full cell is within the same heading row
                    //that way we can filter out other results.
                    if (temp_heading.Count > 1)
                    {
                        MessageBox.Show("Multiple search results for the heading were found." +
                    " Cell address = " + temp_heading[0].Address.ToString() +
                    "Search String = " +heading.headingName);
                        error_found = true;
                        return false;
                    }
                
                }

            }



            return true;
        }


    }
}
