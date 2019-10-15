using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    class ItemWrap
    {
        public  Excel.Range fullCell;

    }
    class DateItemWrap : ItemWrap
    {
        public DateTime content;
    }
    class StrItemWrap: ItemWrap
    {
        public String content;
    }
    class LongItemWrap : ItemWrap
    {
        public long content;
    }

}
