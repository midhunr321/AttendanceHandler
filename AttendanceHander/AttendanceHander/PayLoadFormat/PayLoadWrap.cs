using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander.PayLoadFormat
{
   public class PayLoadWrap
    {
        public List<Day> days;
        public class Day
        {
            public StrItemWrap company;
            public DateItemWrap date;
            public StrItemWrap section;
            public StrItemWrap job;

            public StrItemWrap serialNo;
            public StrItemWrap code;
            public StrItemWrap name; 
            public StrItemWrap design;
            public StrItemWrap job_siteNo;
            public DecimalItemWrap workTime;
            public DecimalItemWrap noBreak;
            public DecimalItemWrap overTime;
           

        }
    }
}
