using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
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

        public void MAIN_understand_the_excel_sheet(out Boolean error_found)
        {
            error_found = false;
            if (SiGlobalVars.Instance.dailyTransHeadings == null)
            {
                SiGlobalVars.Instance.dailyTransHeadings
                    = new DailyTransHelper.Headings();
            }

            if (SiGlobalVars.Instance.dailyTransWraps == null)
                SiGlobalVars.Instance.dailyTransWraps = new List<DailyTransWrap>();

            find_headings(ref SiGlobalVars.
                Instance.dailyTransHeadings, out error_found);
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


        private void read_each_rows_of_data(out Boolean error_occured)
        {
            error_occured = false;
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            //now we have to read the rows
            //our beginning position to start reading the rows will be
            //below the personal Number heading cell

            Excel.Range personnelNo = SiGlobalVars.Instance
                .dailyTransHeadings.personnelNo.fullCell;

           


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
                    ref SiGlobalVars.Instance.dailyTransWraps,
                    out error_occured, out reached_empty_space_area);
                if (error_occured == true)
                    return;
                if (reached_empty_space_area == true)
                    break;

            }

        }

        private void read_row(Excel.Range row,
       ref List<DailyTransWrap> dailyTransWraps,
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

            DailyTransWrap dailyTransWrap = new DailyTransWrap();
            do
            {
                //first nextFullCell is firstFullCell
                //so 
                var currentFullCell = nextFullCell;

                Boolean result1;
                result1 = feed_datas_of_single_row(ref dailyTransWrap,
                    currentFullCell,
                          SiGlobalVars.Instance.dailyTransHeadings,
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


            dailyTransWraps.Add(dailyTransWrap);

        }

        private void feed_time_data_to_dataWrap(ref DateItemWrap time_data,
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
           (ref DailyTransWrap dailyTransWrap,
          Excel.Range fullCell, DailyTransHelper.Headings headings,
          out Boolean error_occured, out Boolean reached_empty_space_or_invalid_data)
        {
            error_occured = false;
            reached_empty_space_or_invalid_data = false;
            if (fullCell.Column > headings.dataFrom.fullCell.Column)
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
                        if (dailyTransWrap.personnelNo == null)
                            dailyTransWrap.personnelNo = new StrItemWrap();
                        dailyTransWrap.personnelNo.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        dailyTransWrap.personnelNo.fullCell = fullCell;
                        dailyTransWrap.personnelNo.heading = heading;
                        return true;
                    }
                    else if (heading.Equals(headings.firstName))
                    {
                        if (dailyTransWrap.firstName == null)
                            dailyTransWrap.firstName = new StrItemWrap();
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
                        dailyTransWrap.firstName.content = extractedName;
                        dailyTransWrap.firstName.fullCell = fullCell;
                        dailyTransWrap.firstName.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.lastName))
                    {
                        if (dailyTransWrap.lastName == null)
                            dailyTransWrap.lastName = new StrItemWrap();

                        dailyTransWrap.lastName.content = eXCEL_HELPER
                            .get_value_of_merge_cell(fullCell);
                        dailyTransWrap.lastName.fullCell = fullCell;
                        dailyTransWrap.lastName.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.position))
                    {
                        if (dailyTransWrap.position == null)
                            dailyTransWrap.position = new StrItemWrap();
                        dailyTransWrap.position.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.position.fullCell = fullCell;
                        dailyTransWrap.position.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.department))
                    {
                        if (dailyTransWrap.department == null)
                            dailyTransWrap.department = new StrItemWrap();
                        dailyTransWrap.department.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.department.fullCell = fullCell;
                        dailyTransWrap.department.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.date))
                    {
                        if (dailyTransWrap.date == null)
                            dailyTransWrap.date = new DateItemWrap();
                        String extractedDate_in_string = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        DateTime extractedDAte;

                        if (DateTime.TryParse(extractedDate_in_string, out extractedDAte)
                        == true)
                            dailyTransWrap.date.content = extractedDAte;
                        else
                        {
                            MessageBox.Show("Found invalid date in cell = " +
                               fullCell.Address);
                            error_occured = true;
                            return false;
                        }
                        dailyTransWrap.date.contentInString =
                            eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.date.fullCell = fullCell;
                        dailyTransWrap.date.heading = heading;

                        //reaching the total over time
                        //as you know after total over time it is overtime datas
                        //so we need to break from this iteration now

                        return true;

                    }
                    else if (heading.Equals(headings.time))
                    {

                        CommonOperations.feed_time_data_to_dataWrap(ref dailyTransWrap.time,
                            eXCEL_HELPER, fullCell, heading, (DateTime)dailyTransWrap.date.content);


                        return true;

                    }
                    else if (heading.Equals(headings.punchStatus))
                    {
                        if (dailyTransWrap.punchStatus == null)
                            dailyTransWrap.punchStatus = new StrItemWrap();
                        dailyTransWrap.punchStatus.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.punchStatus.fullCell = fullCell;
                        dailyTransWrap.punchStatus.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.workCode))
                    {
                        if (dailyTransWrap.workCode == null)
                            dailyTransWrap.workCode = new StrItemWrap();
                        dailyTransWrap.workCode.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.workCode.fullCell = fullCell;
                        dailyTransWrap.workCode.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.gpsLocation))
                    {
                        if (dailyTransWrap.gpsLocation == null)
                            dailyTransWrap.gpsLocation = new StrItemWrap();
                        dailyTransWrap.gpsLocation.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.gpsLocation.fullCell = fullCell;
                        dailyTransWrap.gpsLocation.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.area))
                    {
                        if (dailyTransWrap.area == null)
                            dailyTransWrap.area = new StrItemWrap();
                        dailyTransWrap.area.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.area.fullCell = fullCell;
                        dailyTransWrap.area.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.deviceName))
                    {
                        if (dailyTransWrap.deviceName == null)
                            dailyTransWrap.deviceName = new StrItemWrap();
                        dailyTransWrap.deviceName.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.deviceName.fullCell = fullCell;
                        dailyTransWrap.deviceName.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.deviceSerialNo))
                    {
                        if (dailyTransWrap.deviceSerialNo == null)
                            dailyTransWrap.deviceSerialNo = new StrItemWrap();
                        dailyTransWrap.deviceSerialNo.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.deviceSerialNo.fullCell = fullCell;
                        dailyTransWrap.deviceSerialNo.heading = heading;

                        return true;
                    }
                    else if (heading.Equals(headings.dataFrom))
                    {
                        if (dailyTransWrap.dataFrom == null)
                            dailyTransWrap.dataFrom = new StrItemWrap();
                        dailyTransWrap.dataFrom.content = eXCEL_HELPER.get_value_of_merge_cell(fullCell);
                        dailyTransWrap.dataFrom.fullCell = fullCell;
                        dailyTransWrap.dataFrom.heading = heading;

                        return true;
                    }
                }
            }
            return false;
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



            return true;
        }
    }
}
