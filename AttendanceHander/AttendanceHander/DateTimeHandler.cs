using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
  public  class DateTimeHandler
    {
        public static DateTime mix_different_date_and_time(DateTime new_date,
          DateTime time)
        {
            return (new_date.Date + time.TimeOfDay);
        }

    }
}
