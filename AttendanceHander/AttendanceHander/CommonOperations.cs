using System;
using System.Collections.Generic;
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

        public CommonOperations(Excel.Worksheet worksheet)
        {
            this.worksheet = worksheet;
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
