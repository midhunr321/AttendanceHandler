using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
   public class MepStyleWrap
    {
        public class StrItemWrap : AttendanceHander.StrItemWrap { }
        public class DateOvertime : AttendanceHander.DateOvertime { }

        public StrItemWrap serialNo;
            public StrItemWrap code; //Employee no.
            public StrItemWrap name;
            public StrItemWrap designation;
            public StrItemWrap siteNo;
            public StrItemWrap totalOvertime;
            public List<DateOvertime> dateOvertimes;
        
 
    }
}
