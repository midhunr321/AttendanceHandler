using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander
{
    class DailyTransWrap
    {
        ItemWrap personalNo;
        String firstName;
        String position;
        String department;
        DateTime date;
        String area;
        String deviceName;
        DateTime totalTimeWorked;
    }
}
