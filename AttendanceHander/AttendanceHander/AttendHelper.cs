using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{

    class AttendHelper
    {
        public Excel.Worksheet current_worksheet;
        public Excel.Application excel_app;
        EXCEL_HELPER eXCEL_HELPER;
        int total_used_row;
        int total_used_col;
    }
}
    