using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander.DailyTransactions
{
   public class DailyTransWrap
    {
        public StrItemWrap personnelNo; //Personnel No.
        public StrItemWrap firstName; //First Name - both are same
        public StrItemWrap lastName; //LastName is normally not occupied
        public StrItemWrap position;
        public StrItemWrap department;
        public DateItemWrap date;
        public DateItemWrap time;
        public StrItemWrap punchStatus;
        public StrItemWrap workCode;
        public StrItemWrap gpsLocation;
        public StrItemWrap area;
        public StrItemWrap deviceName;
        public StrItemWrap deviceSerialNo;
        public StrItemWrap dataFrom;
    }
}
