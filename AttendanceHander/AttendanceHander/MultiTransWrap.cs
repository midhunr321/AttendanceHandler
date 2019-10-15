using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    class MultiTransWrap
    {
        Dictionary<Excel.Range, long> personalNo;
        Dictionary<Excel.Range, String> firstName;
        Dictionary<Excel.Range, String> position;
        Dictionary<Excel.Range, String> department;
       Dictionary<Excel.Range,DateTime> checkIn;
        Dictionary<Excel.Range,DateTime> checkOut;
        Dictionary<Excel.Range, DateTime> date;
        Dictionary<Excel.Range, DateTime> totalTimeWorked;

    }
}
