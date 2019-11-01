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

        public static Boolean Compare_dates_only(DateTime dateTime1, DateTime dateTime2)
        {
            
            if (dateTime1.Year == dateTime2.Year &&
                dateTime1.Month == dateTime2.Month &&
                dateTime1.Day == dateTime2.Day)
                return true;
            else
                return false;
        }

    }
}
