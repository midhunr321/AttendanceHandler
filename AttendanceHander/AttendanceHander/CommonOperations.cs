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
            else if (count == 0)
            {
                //that means there is no 'S'
                // means invalid site no

                String error_message = "No 'S' was detected in the site no";
                MessageBox.Show(error_message + " Cell Address = " +
                        deviceName_in_dailyTransFormat.fullCell.Address
                        + " Content = " + deviceName_in_dailyTransFormat.content);
                return null;
            }
            //first trim free spaces
            sourceString = sourceString.Trim();
            //now check if the site no begins with 'S'
            //ie the 'S' should be at the very beginning.
            if (sourceString.IndexOf('S', 0) > 0)
            {
                //i.e the char 'S' is not at the very beginning
                //which means error
                String error_message = "Char 'S' was not at the very beginning";
                MessageBox.Show(error_message + " Cell Address = " +
                        deviceName_in_dailyTransFormat.fullCell.Address
                        + " Content = " + deviceName_in_dailyTransFormat.content);
                return null;
            }
            return sourceString.Replace('S', 'M');
        }

        internal static void Close_running_workbook(Excel.Workbook workbook)
        {
            List<Excel.Workbook> workbooks =
                            new List<Excel.Workbook>() { workbook };
            //if excel is open, close it
            if (EXCEL_HELPER
                .Given_workbooks_are_running_in_background
                (workbooks))
                workbook.Close();
        }

        public static String replace_first_S_in_siteNo_with_M(StrItemWrap deviceName_in_dailyTransFormat)
        {
            string siteno = deviceName_in_dailyTransFormat.content.Trim();
            //after trim the first char should be 'S'
            Char[] array = siteno.ToCharArray();
            if (array.Length == 0)//if null then return the same
                return siteno;

            if (array[0] == 'S')
                array[0] = 'M';
            siteno = new string(array);
            siteno = siteno.Trim();
            return siteno;
        }
        public static String convert_siteNo_to_SiteNoMechFormat_ShortName(StrItemWrap deviceName_in_dailyTransFormat)
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
        public static Boolean compare_multiTrans_employeeNo_to_MepAndPayLoad_employeeNo
            (String mepOrPayLoad_employeeNo, String multiTrans_employeeNo)
        {
            //String mep_Trimmed = mepStyle_employeeNo.Trim('/');
            String mep_Trimmed = mepOrPayLoad_employeeNo
                .Substring(mepOrPayLoad_employeeNo.IndexOf('/') + 1);

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
            foreach (Excel.Range result in searchResults)
            {
                if (result.Row == adjacentHeadingCell1.Row &&
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
            else if (filtered_search_result.Count == 0)
            {
                StackTrace stackTrace = new StackTrace();
                Console.WriteLine(stackTrace.ToString());
                MessageBox.Show("After Search filteration the Search Result became zero");
                return null;
            }
            return filtered_search_result[0];
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

        internal static bool Given_short_siteNo_is_valid(string shortSiteNo)
        {
            StringHandler stringHandler = new StringHandler();
            if (stringHandler
                  .is_this_string_alpha_numeric_or_numeric_or_alpha_only(shortSiteNo)
                  != All_const.str_type.Alphanumeric)
                return false;

            //TODO in future we need to check if Capital M is available or not

            return true;

        }
        private static void Refresh_holiday_dates_display(Label label_holidays)
        {
            if (SiGlobalVars.Instance.Holidays == null)
            {
                label_holidays.Text = "Holidays : Null";
                return;
            }
            label_holidays.Text = "Holidays : ";
            foreach (var holiday in SiGlobalVars.Instance.Holidays)
            {
                label_holidays.Text = label_holidays.Text +
                    holiday.ToShortDateString() + ", ";
            }

        }

        internal static Boolean Display_holiday_selectorForm(Form previousForm,
            Label label_holidays)
        {
            Boolean displayHolidaySelectorDialog = true;
            if (SiGlobalVars.Instance.Holidays != null)
            {
                displayHolidaySelectorDialog = false;

                var messageResult =
                    MessageBox.Show("It seems like you have already selected the holidays." +
                    " Would you like to Re-select?", "Re-select Holidays?", MessageBoxButtons.YesNo);

                if (messageResult == DialogResult.Yes)
                    displayHolidaySelectorDialog = true;

            }

            if (displayHolidaySelectorDialog == true)
            {
                Form_holidaysSelector form_HolidaysSelector = new Form_holidaysSelector(previousForm);
                DialogResult dialogResult = form_HolidaysSelector.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    var holidays = form_HolidaysSelector.SelectedHolidays;
                    if (SiGlobalVars.Instance.Holidays == null)
                        SiGlobalVars.Instance.Holidays = new List<DateTime>();
                    SiGlobalVars.Instance.Holidays = holidays;

                    Refresh_holiday_dates_display(label_holidays);

                    SiGlobalVars.Instance.Holidays = holidays;

                    return true;
                }
            }


            return false;
        }
    }
}
