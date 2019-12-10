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

        private void extract_data_from_sheets(out Boolean error_found)
        {
            error_found = false;
            int totalDaysInMonth = get_total_days_in_this_month();
            if(totalDaysInMonth <0)
            {
                error_found = true;
                return;
            }

            if (loop_through_each_date_sheets_in_order(totalDaysInMonth)
                 == false)
            {
                error_found = true;
                return;
            }



            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
            }


        }

        private bool loop_through_each_date_sheets_in_order(int totalDaysinMonth,
            out Boolean error_found)
        {
            error_found = false;
            //first find the first sheet.

            Excel.Worksheet firstSheet=null;
  
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


            for(int day = 2; day <= totalDaysinMonth; day++)
            {
                //now check if the sheets are in correct order
                Excel.Worksheet nextSheet = firstSheet.Next();
                if(nextSheet.Name.Trim() != day.ToString())
                {
                    MessageBox.Show("Sheets might not be in order; Expected sheet = "
                        + day.ToString() + "; but the sheet obtained = " + nextSheet.Name);
                    error_found = true;
                    return false;
                }


                find_headings(ref SiGlobalVars.
                 Instance.payLoadHeadings, out error_found);
                
                if (error_found == true)
                return false;



                //now we got all the headings
                //connect heading and datas together

                //now that we got all headings
                //we need to start with the rows

                read_each_rows_of_data(out error_found);
                if (error_found == true)
                    return false;

            }


            return true; 
        }

        private void read_each_rows_of_data(out bool error_occured)
        {
            error_occured = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            //now we have to read the rows
            //our beginning position to start reading the rows will be
            //below the personal Number heading cell

            Excel.Range serialNoHeading = SiGlobalVars.Instance
                .payLoadHeadings.serialNo.fullCell;

            int lastHeadingRow = EXCEL_HELPER.get_last_row_no_of_a_merged_cell(serialNoHeading);
            //now go to below adjacent cell to personnal no.
            int firstDataRowNo = lastHeadingRow + 1;



            foreach (Excel.Range row in worksheet.UsedRange.Rows)
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
                    ref SiGlobalVars.Instance.payloadWraps,
                    out error_occured, out reached_empty_space_area);
                if (error_occured == true)
                    return;
                if (reached_empty_space_area == true)
                    break;

            }

        }

        private void read_row(Excel.Range row, 
            ref List<PayLoadWrap> payloadWraps, 
            out bool error_occured, out bool reached_empty_space_area, 
            Excel.Worksheet sheet)
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

            PayLoadFormat.PayLoadWrap payLoadWrap = new PayLoadWrap();
            do
            {
                //first nextFullCell is firstFullCell
                //so 
                var currentFullCell = nextFullCell;

                Boolean result1;
                result1 = feed_datas_of_single_row(ref payLoadWrap,
                    currentFullCell,
                          SiGlobalVars.Instance.payLoadHeadings,
                         out error_occured, out reached_empty_space_area, sheet);

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


            payloadWraps.Add(payLoadWrap);
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

            if(assumedDate==null)
            {
                MessageBox.Show("Date in Multiple Transaction is null; Cell Address = "
                    + SiGlobalVars.Instance
                .multiTransWraps.First().date.fullCell.Address +
                "; Content = ");
                return -1;
            }
            int noOfDays = DateTime.DaysInMonth(assumedDate.Year,assumedDate.Month);


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

            if (SiGlobalVars.Instance.payloadWraps == null)
                SiGlobalVars.Instance.payloadWraps 
                    = new List<PayLoadWrap>();


            extract_data_from_sheets(out error_found);

         



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
