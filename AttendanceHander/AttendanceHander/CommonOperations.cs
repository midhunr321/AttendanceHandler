using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
   public class CommonOperations
    {
        private Excel.Worksheet worksheet;


        public static Boolean compare_multiTrans_employeeNo_to_MepStyle_employeeNo
            (String mepStyle_employeeNo, String multiTrans_employeeNo )
        {
            //String mep_Trimmed = mepStyle_employeeNo.Trim('/');
            String mep_Trimmed = mepStyle_employeeNo
                .Substring(mepStyle_employeeNo.IndexOf('/') + 1);

            mep_Trimmed = mep_Trimmed.Trim('0');
            //eg
            //before = 02/04532
            //after trim = 4532

            mep_Trimmed = mep_Trimmed.Trim(); //trim unwanted white space

            String multiTrans_trimmed = multiTrans_employeeNo.Trim();  //remove white spaces

            multiTrans_trimmed = multiTrans_trimmed.TrimStart(new char[] { '0' });

            if (mep_Trimmed == multiTrans_trimmed)
                return true;
            else
                return false;

        }
        public CommonOperations(Excel.Worksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        
        public static Excel.Range filter_searchResult_by_comparing_row_no_of_adjacent_headings
            (List<Excel.Range> searchResults,
            Excel.Range adjacentHeadingCell1, Excel.Range adjacentHeadingCell2)
        {
            Boolean all_rows_are_equal = false;
            //if all are in same row means they are in the heading row

            List<Excel.Range> filtered_search_result = new List<Excel.Range>();
            foreach(Excel.Range result in searchResults)
            {
                if(result.Row == adjacentHeadingCell1.Row &&
                    result.Row == adjacentHeadingCell2.Row)
                {
                    filtered_search_result.Add(result);

                }
            }

            if (filtered_search_result.Count > 1)
            {
                StackTrace stackTrace = new StackTrace();
                Console.WriteLine(stackTrace.ToString());
                MessageBox.Show("Multiple search results even after filtering in the same row");
                return null;
            }
           else if (filtered_search_result.Count ==0)
            {
                StackTrace stackTrace = new StackTrace();
                Console.WriteLine(stackTrace.ToString());
                MessageBox.Show("After Search filteration the Search Result became zero");
                return null;
            }
            return filtered_search_result[0];
        }

        public static void feed_time_data_to_dataWrap(ref TimeSpanItemWrap time_data,
      EXCEL_HELPER eXCEL_HELPER, Excel.Range fullCell,
      HeadingWrap heading, DateTime date_of_time)
        {
            //that is employee no
            if (time_data == null)
                time_data = new TimeSpanItemWrap();
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
        public static void feed_time_data_to_dataWrap(ref DateItemWrap time_data,
        EXCEL_HELPER eXCEL_HELPER, Excel.Range fullCell,
        HeadingWrap heading, DateTime date_of_time)
        {
            //that is employee no
            if (time_data == null)
                time_data = new DateItemWrap();
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

        public static bool employeeNo_is_valid(string extractedEmployeeNo)
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
        public static Boolean name_is_valid(String name)
        {
            if (name == null)
                return false;
            if (String.IsNullOrEmpty(name))
                return false;
            if (String.IsNullOrWhiteSpace(name))
                return false;
            String trimmedName = name.Trim();
            //first remove the unwanted white space from the beginning and the ending
            //using the trim 
            //then check if the string is empty or not
            StringHandler stringHandler = new StringHandler();
            if (trimmedName.Length < 3)
                return false;
            if (stringHandler.is_this_string_alpha_numeric_or_numeric_or_alpha_only(trimmedName)
                == All_const.str_type.Numeric)
                return false;

            return true;

        }

    }
}
