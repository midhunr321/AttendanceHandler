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
            public HeadingWrap mepStyleHeading =
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
                {mepStyleHeading,serialNo,
                    code,name,designation,siteNO,
                    totalOvertime,date }.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
        public void understand_the_excel_sheet()
        {
            if (SiGlobalVars.Instance.mepStyleHeadings == null)
            {
                SiGlobalVars.Instance.mepStyleHeadings = new Headings();
            }

            find_headings_except_overtimeDates(ref SiGlobalVars.
                Instance.mepStyleHeadings);//if all the headings
            //are found, it means the opened excel file is Mep style 
            //plumbers time sheet
            understand_the_month_and_year_of_the_sheet();
            find_overtime_dates_headings();
            //now that we got all headings
            //we need to start with the rows
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
                SiGlobalVars.Instance.mepStyleWraps = new MepStyleWrap();

            if (SiGlobalVars.Instance.mepStyleWraps.
                timesheetDate == null)
            {
                SiGlobalVars.Instance.
                    mepStyleWraps.timesheetDate = new DateTime();
            }

            SiGlobalVars.Instance.
                mepStyleWraps.timesheetDate =
                DateTime.Parse(timesheetDate);

        }



        private void find_overtime_dates_headings()
        {
            var timesheetDate = SiGlobalVars.Instance
                .mepStyleWraps.timesheetDate;
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

                lastDay = nextday;
            }

        }

        private Boolean find_headings_except_overtimeDates(ref MepStyleHelper.Headings headings)
        {
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            foreach (HeadingWrap heading in headings)
            {
                List<Excel.Range> temp_heading = new List<Excel.Range>();
                temp_heading =
                    eXCEL_HELPER.find_fix_column_heading(heading.headingName,
                    Excel.XlSearchDirection.xlNext,
                    Excel.XlSearchOrder.xlByRows,false);

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

                }

            }



            return true;
        }


    }
}
