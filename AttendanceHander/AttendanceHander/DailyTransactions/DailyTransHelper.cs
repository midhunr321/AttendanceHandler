using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AttendanceHander.MultipleTransaction;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander.DailyTransactions
{
   public class DailyTransHelper
    {
        private Excel.Worksheet worksheet;
        private Excel.Workbook workbook;

        public DailyTransHelper(Excel.Worksheet worksheet, Excel.Workbook workbook)
        {
            this.worksheet = worksheet;
            this.workbook = workbook;
        }

        public class Headings : IEnumerable<HeadingWrap>
        {
            public HeadingWrap sheetHeading =
                new HeadingWrap("Transactions");
            public HeadingWrap personnelNo =
                new HeadingWrap("Personnel No.");
            public HeadingWrap firstName = new HeadingWrap("First Name");
            public HeadingWrap lastName = new HeadingWrap("Last Name");
            public HeadingWrap position = new HeadingWrap("Position");
            public HeadingWrap department = new HeadingWrap("Department");
            public HeadingWrap date = new HeadingWrap("Date");
            public HeadingWrap time = new HeadingWrap("Time");
            public HeadingWrap punchStatus = new HeadingWrap("Punch Status");
            public HeadingWrap workCode = new HeadingWrap("Work Code");
            public HeadingWrap gpsLocation = new HeadingWrap("GPS Location");
            public HeadingWrap area = new HeadingWrap("Area");
            public HeadingWrap deviceName = new HeadingWrap("Device Name");
            public HeadingWrap deviceSerialNo = new HeadingWrap("Device SN");
            public HeadingWrap dataFrom = new HeadingWrap("Data From");

            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {sheetHeading,personnelNo,
                    firstName,lastName,position,department,
                    date,time,punchStatus,workCode, gpsLocation,
                area,deviceName, deviceSerialNo,dataFrom  }.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }


        }

        public Boolean MAIN_understand_the_excel_sheet()
        {
            Boolean error_found = false;
            if (SiGlobalVars.Instance.multiTransHeadings == null)
            {
                SiGlobalVars.Instance.dailyTransHeadings
                    = new DailyTransHelper.Headings();
            }


            find_headings(ref SiGlobalVars.
                Instance.dailyTransHeadings, out error_found);
            if (error_found == true)
                return false;


            //now we got all the headings
            //connect heading and datas together

            //now that we got all headings
            //we need to start with the rows

            read_each_rows_of_data(out error_found);
            if (error_found == true)
                return false;

            return true;
        }


        private void read_each_rows_of_data(out Boolean error_occured)
        {
            error_occured = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            //now we have to read the rows
            //our beginning position to start reading the rows will be
            //below the personal Number heading cell

            Excel.Range personnelNo = SiGlobalVars.Instance
                .dailyTransHeadings.personnelNo.fullCell;

            //Excel.Range personnelNo = SiGlobalVars.Instance
            //    .multiTransHeadings.personnelNo.fullCell;


            //now go to below adjacent cell to personnal no.
            Excel.Range firstDataRowCell =
                eXCEL_HELPER
                .return_immediate_below_cell(personnelNo);

            foreach (Excel.Range row in worksheet.UsedRange.Rows)
            {
                //our first data row starts from firstDataRowCell
                //so skip the rows above (which are headings)
                if (row.Row < firstDataRowCell.Row)
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
        private Boolean find_headings(ref DailyTransHelper.Headings headingWraps,
           out Boolean error_found)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            foreach (HeadingWrap heading in headingWraps)
            {

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
                    
                        //if it is not special case
                        //and if two search results for heading means something is wrong
                        MessageBox.Show("Multiple search results were found for: Cell = " +
                            search_results[0].Address.ToString() + "Content = "+
                            heading.headingName);
                        error_found = true;
                        return false;
                    
                }

            }



            return true;
        }
    }
}
