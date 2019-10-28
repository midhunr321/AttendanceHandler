using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    public class MultiTransWrap
    {
        LongItemWrap personalNo;
        StrItemWrap firstName;
        StrItemWrap designation;
        StrItemWrap department;
        DateItemWrap checkIn;
        DateItemWrap checkOut;
        DateItemWrap date;
        DateItemWrap totalTimeWorked;

    }
}
