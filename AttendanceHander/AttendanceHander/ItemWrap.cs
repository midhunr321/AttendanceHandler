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
        public Nullable<DateTime> content;
        public HeadingWrap heading;
        
        
    }
    public class StrItemWrap : ItemWrap
    {
        public String content;
        public HeadingWrap heading;

    }
    public class LongItemWrap : ItemWrap
    {
        public long content;
        public HeadingWrap heading;

    }
    public class DateOvertime : ItemWrap
    {
        public String date_day;
        public HeadingWrap heading;
        public DateTime date;
        public String siteNo;//for each overtime there might be a site no
        public String overtime;
    }

}
