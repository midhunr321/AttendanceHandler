using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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

      
     
        public static void modify_value_in_cell(Excel.Range fullCell, String Value,
            Color color)
        {
            fullCell.Value = Value;
            fullCell.Font.Color = color;
            fullCell.Font.Italic = true;
            fullCell.Font.Bold = true;
            
        }
        private static String replace_S_in_string_with_M(String sourceString,
            StrItemWrap deviceName_in_dailyTransFormat)
        {
           
            int count = sourceString.Count(x => x == 'S');
            if (count > 1)
            {
                //that means more than one S is there
                //normally it shouldn't happen
                //a healthy one should look like this S223
                //so that means an error

                String error_message = "More than one 'S' character was found in the source string";
                MessageBox.Show(error_message + " Cell Address = " +
                        deviceName_in_dailyTransFormat.fullCell.Address
                        + " Content = " + deviceName_in_dailyTransFormat.content);
                return null;
            }
             return sourceString.Replace('S', 'M');
        }
        public static String convert_siteNo_to_SiteNoMechFormat(StrItemWrap deviceName_in_dailyTransFormat)
        {
            //note that deviceName is the site no in daily transactions
            //first just trim the empty spaces
            String siteNo = deviceName_in_dailyTransFormat.content.Trim();

            //check how many '-' is there in the string
            //eg S276-1101 
            // in the above, only one '-' is there
            //but if more than one '-', it should throw an error

            int count = siteNo.Count(x => x == '-');
            if (count > 1)
            {
                //means '-' repeats more than one time
                //which means error

                MessageBox.Show("SiteNo or Device name in Daily Transactions is invalid; it has more than one '-'" +
                    " Content=" + deviceName_in_dailyTransFormat.content + " Cell address" + deviceName_in_dailyTransFormat.fullCell.Address);
                return null;
            }
            else if (count == 0)
            {
                //means example if like this S223
                //then no need to filter any thing
                //only change 'S' to 'M'
             
                siteNo = replace_S_in_string_with_M(siteNo,
                    deviceName_in_dailyTransFormat);
                if (siteNo == null)
                    return null;
               
            }
            else if (count == 1)
            {
                //means example if like this S269-D2
                //in this case get all char till '-'
                //so split it by '-'
                siteNo = siteNo.Split('-')[0];
                //now trim unwanted spaces
                siteNo = siteNo.Trim();
                //now replace s with M
                siteNo = replace_S_in_string_with_M(siteNo, 
                    deviceName_in_dailyTransFormat);
            }

            return siteNo;
        }
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
