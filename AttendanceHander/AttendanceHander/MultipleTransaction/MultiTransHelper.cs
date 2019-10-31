using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
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

        public Boolean MAIN_understand_the_excel_sheet()
        {
            Boolean error_found = false;
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
                      if(  filterout_multiple_search_results_for_special_case_and_assign_values
                        (search_results, heading)
                            ==false)
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

                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.checkInTime1,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);


                        return true;

                    }
                    else if (heading.Equals(headings.checkOutTime1))
                    {
                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.checkOutTime1,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.workingTime1))
                    {
                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.workingTime1,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.checkInTime2))
                    {
                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.checkInTime2,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.checkOutTime2))
                    {
                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.checkOutTime2,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.workingTime2))
                    {
                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.workingTime2,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }
                    else if (heading.Equals(headings.totalTimeWorked))
                    {
                        CommonOperations.feed_time_data_to_dataWrap(ref multiTransWrap.totalTimeWorked,
                            eXCEL_HELPER, fullCell, heading, (DateTime)multiTransWrap.date.content);



                        return true;

                    }


                }
            }
            return false;
        }



    }
}
