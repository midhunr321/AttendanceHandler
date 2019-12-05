using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander.MultipleTransaction
{
   public class MultiTransWrap
    {
        public StrItemWrap personnelNo; //Personnel No.
        public StrItemWrap firstName; //First Name - both are same
        public StrItemWrap lastName; //LastName is normally not occupied
        public StrItemWrap position;
        public StrItemWrap department;
        public DateItemWrap date;
        public DateItemWrap checkInTime1;
        public DateItemWrap checkOutTime1;
        public TimeSpanItemWrap workingTime1;
        public DateItemWrap checkInTime2;
        public DateItemWrap checkOutTime2;
        public TimeSpanItemWrap workingTime2;
        public TimeSpanItemWrap totalTimeWorked;
        public StrItemWrap siteNo;
        public SiteNoMechFormat siteNoMechFormat;


        public class SiteNoMechFormat
        {
            public StrItemWrap shortName;
            public StrItemWrap fullName;
        }
    }

}
