using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    public class ItemWrap
    {
        public Excel.Range fullCell;

    }
    public class DateItemWrap : ItemWrap
    {
        public DateTime content;
    }
    public class StrItemWrap : ItemWrap
    {
        public String content;
    }
    public class LongItemWrap : ItemWrap
    {
        public long content;
    }
    public class DateOvertime : ItemWrap
    {
        public int date_day;
        public DateTime date;
        public int overtime;
    }

}
