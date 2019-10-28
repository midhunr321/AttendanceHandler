using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            public HeadingWrap multiTransHeading =
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

            public Dictionary<int, HeadingWrap> overtimeDays;

            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {multiTransHeading,personnelNo,firstName,lastName
                ,position,department,date,checkInTime1,checkOutTime1,workingTime1,
                checkInTime2,checkOutTime2,workingTime2,
                totalTimeWorked}.GetEnumerator());
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
                SiGlobalVars.Instance.multiTransHeadings
                    = new MultipleTransaction.MultiTransHelper.MultiHeadings();
            }


            find_headings(ref SiGlobalVars.
                Instance.multiTransHeadings, out error_found);
            if (error_found == true)
                return false;


            //now we got all the headings
            //connect heading and datas together

            //now that we got all headings
            //we need to start with the rows

            read_each_rows_of_data(out error_ocurred);
            if (error_ocurred == true)
                return false;

            return true;
        }
        private Boolean find_headings(ref MultipleTransaction.MultiTransHelper
            .MultiHeadings headingWraps,
            out Boolean error_found)
        {
            error_found = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            foreach (HeadingWrap heading in headingWraps)
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
                    error_found = true;
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
            ///atleast_one_result_was_true_in_this_row is used
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
                          SiGlobalVars.Instance.mepStyleHeadings,
                         out error_occured, out reached_empty_space_area);

                if (reached_empty_space_area == true)
                    return;


                if (error_occured == true)
                    return;

                //Boolean result2;

              

                //now get the site transfer start and end dates
                Boolean stop_this_row_iteration = false;

                //Boolean result3;
                //result3 = feed_site_transfer_data_of_a_cell
                //      (ref multiTransWrap.dateOvertimes, currentFullCell,
                //         SiGlobalVars.Instance.mepStyleHeadings,
                //        out stop_this_row_iteration);

                //if (stop_this_row_iteration == true)
                //    break;


                //if (nextCell == null)
                //    nextCell = firstFullCell.Next;
                //else
                //    nextCell = nextCell.Next;
                //nextFullCell = eXCEL_HELPER.return_full_merg_cell(nextCell);

                //i++;
            } while (i <= totalNoUsedColumns);






            multiTransWraps.Add(multiTransWrap);

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
                if (heading.Equals(headings.multiTransHeading))
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

                        if (TimeSheetOperations.employeeNo_is_valid(extractedEmployeeNo)
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
                            multiTransWrap.personnelNo = new StrItemWrap();
                        multiTransWrap.personnelNo.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        multiTransWrap.personnelNo.fullCell = fullCell;
                        multiTransWrap.personnelNo.heading = heading;
                        return true;
                    }
                    else if (heading.Equals(headings.firstName))
                    {
                        if (multiTransWrap.firstName == null)
                            multiTransWrap.firstName = new StrItemWrap();
                        String extractedName = eXCEL_HELPER.get_value_of_merge_cell(fullCell);

                        if (TimeSheetOperations.name_is_valid(extractedName)
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
                            multiTransWrap.lastName = new StrItemWrap();

                        multiTransWrap.lastName.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        multiTransWrap.lastName.fullCell = fullCell;
                        multiTransWrap.lastName.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.position))
                    {
                        if (multiTransWrap.position == null)
                            multiTransWrap.position = new StrItemWrap();
                        multiTransWrap.position.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        multiTransWrap.position.fullCell = fullCell;
                        multiTransWrap.position.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.department))
                    {
                        if (multiTransWrap.department == null)
                            multiTransWrap.department = new StrItemWrap();
                        multiTransWrap.department.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        multiTransWrap.department.fullCell = fullCell;
                        multiTransWrap.department.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.date))
                    {
                        if (multiTransWrap.date == null)
                            multiTransWrap.date = new DateItemWrap();
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
                        
                        multiTransWrap.date.fullCell = fullCell;
                        multiTransWrap.date.heading = heading;

                        //reaching the total over time
                        //as you know after total over time it is overtime datas
                        //so we need to break from this iteration now

                        return true;

                    }
                    else if (heading.Equals(headings.checkInTime1))
                    {

                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.checkInTime1,
                            eXCEL_HELPER, fullCell, heading);


                        return true;

                    }
                    else if (heading.Equals(headings.checkOutTime1))
                    {
                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.checkOutTime1,
                            eXCEL_HELPER, fullCell, heading);



                        return true;

                    }
                    else if (heading.Equals(headings.workingTime1))
                    {
                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.workingTime1,
                            eXCEL_HELPER, fullCell, heading);



                        return true;

                    }
                    else if (heading.Equals(headings.checkInTime2))
                    {
                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.checkInTime2,
                            eXCEL_HELPER, fullCell, heading);



                        return true;

                    }
                    else if (heading.Equals(headings.checkOutTime2))
                    {
                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.checkOutTime2,
                            eXCEL_HELPER, fullCell, heading);



                        return true;

                    }
                    else if (heading.Equals(headings.workingTime2))
                    {
                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.workingTime2,
                            eXCEL_HELPER, fullCell, heading);



                        return true;

                    }
                    else if (heading.Equals(headings.totalTimeWorked))
                    {
                        feed_checkIn_or_checkOut_time_to_dataWrap(ref multiTransWrap.totalTimeWorked,
                            eXCEL_HELPER, fullCell, heading);



                        return true;

                    }


                }
            }
            return false;
        }

        private void feed_checkIn_or_checkOut_time_to_dataWrap(ref DateItemWrap checkIn_or_checkOut,
            EXCEL_HELPER eXCEL_HELPER, Excel.Range fullCell, 
            HeadingWrap heading)
        {
            //that is employee no
            if (checkIn_or_checkOut == null)
                checkIn_or_checkOut = new DateItemWrap();
            String extractedDate_in_string
                = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
            DateTime result;

            if (DateTime.TryParseExact(extractedDate_in_string,
                "HH:mm", CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AdjustToUniversal,
                out result)
            == true)
                checkIn_or_checkOut.content = result;
            else
                checkIn_or_checkOut.content = null;

            checkIn_or_checkOut.fullCell = fullCell;
            checkIn_or_checkOut.heading = heading;

            //reaching the total over time
            //as you know after total over time it is overtime datas
            //so we need to break from this iteration now
        }
    }
}
