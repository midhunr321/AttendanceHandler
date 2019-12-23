using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
    class TimeSpanHelper
    {
        internal static bool GivenTimeSpan_is_zeroOrNull(TimeSpan content)
        {
            if (content == null)
                return true;

            var compResult = TimeSpan.Compare(content, TimeSpan.Zero);
            if (compResult == 0 || compResult == -1)
                return true;


            return false;
        }
    }
}
