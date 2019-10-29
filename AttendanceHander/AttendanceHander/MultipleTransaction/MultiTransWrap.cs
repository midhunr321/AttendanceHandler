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
        public DateItemWrap workingTime1;
        public DateItemWrap checkInTime2;
        public DateItemWrap checkOutTime2;
        public DateItemWrap workingTime2;
        public DateItemWrap totalTimeWorked;
        public StrItemWrap siteNo;
    }
}
