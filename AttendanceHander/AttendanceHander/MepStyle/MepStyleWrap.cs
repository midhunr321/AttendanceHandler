using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
   public class MepStyleWrap
    {
        public DateTime timesheetDate;
        public List<DataRow> dataRows;
       public class DataRow
        {
            public LongItemWrap serialNo;
            public LongItemWrap code; //Employee no.
            public StrItemWrap name;
            public StrItemWrap designation;
            public StrItemWrap siteNo;
            public LongItemWrap totalOvertime;
            public List<DateOvertime> dateOvertimes;
        }
 
    }
}
