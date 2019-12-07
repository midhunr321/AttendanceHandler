using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander.PayLoadFormat
{
  public class PayLoadHelper
    {
        private Excel.Worksheet worksheet;
        private Excel.Workbook workbook;

        public PayLoadHelper(Excel.Worksheet worksheet, Excel.Workbook workbook)
        {
            this.worksheet = worksheet;
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

        private void loop_through_sheets(out Boolean error_found)
        {
            error_found = false;
            List<String> expected_sheetNames = get_sheetNames_based_on_month();
            if(expected_sheetNames ==null)
            {
                error_found = true;
                return;
            }

            if (all_expectedSheets_are_available(expected_sheetNames)
                 == false)
            {
                error_found = true;
                return;
            }



            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                keyValuePairs.Add(sheet.Name, sheet);
            }


            find_headings(ref SiGlobalVars.
             Instance.payLoadHeadings, out error_found);
            if (error_found == true)
                return;

            //now we got all the headings
            //connect heading and datas together

            //now that we got all headings
            //we need to start with the rows
        }

        private bool all_expectedSheets_are_available(List<string> expected_sheetNames)
        {
            foreach(var expectedName in expected_sheetNames)
            {
                Boolean this_sheet_is_available =false;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name.Trim() == expectedName.Trim())
                        this_sheet_is_available = true;
                }

                if (this_sheet_is_available == false)
                {
                    MessageBox.Show("Couldn't find sheet no = " + expectedName + " in PayloadFormat");
                    return false;
                }
            }

            return true; 
        }

        private List<string> get_sheetNames_based_on_month()
        {
            //no of sheets depends upon number of days in a month
            //eg febraury = 28 
            //so first we need to understand which month and year
            //we can refer it from the multitranswrap
            //assume any date from and employee and thus we can get the date and month
            DateTime assumedDate = SiGlobalVars.Instance
                .multiTransWraps.First().date.content.Value;

            if(assumedDate==null)
            {
                MessageBox.Show("Date in Multiple Transaction is null; Cell Address = "
                    + SiGlobalVars.Instance
                .multiTransWraps.First().date.fullCell.Address +
                "; Content = ");
                return null;
            }
            int noOfDays = DateTime.DaysInMonth(assumedDate.Year,assumedDate.Month);

            List<String> days = new List<string>();
            for (int i= 1; i <= noOfDays; i++){
                days.Add(i.ToString());
            }

            return days;
        }

        public void MAIN_understand_the_excel_sheet(out Boolean error_found)
        {
            error_found = false;
            if (SiGlobalVars.Instance.payLoadHeadings == null)
            {
                SiGlobalVars.Instance.payLoadHeadings
                    = new PayLoadHelper.PayloadHeadings();
            }

            if (SiGlobalVars.Instance.payloadWraps == null)
                SiGlobalVars.Instance.payloadWraps 
                    = new List<PayLoadWrap>();


            loop_through_sheets();

         



        }

        private Boolean find_headings(ref PayloadHeadings payLoadHeadings,
            out bool error_found)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
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
