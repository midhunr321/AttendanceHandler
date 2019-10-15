using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
   public static class All_const
    {

        public enum timesheet_type
        {
            Plumbers = 0,
            Electrical = 1,
            None = 3
        };

        public enum str_type
        {
            Alphanumeric,
            Numeric,
            Alpha_only,
            Null
        };
        public enum search_position
        {
            Top,
            Bottom,
            Null
        };
       public enum flowsheet_type {
            Drawing = 0,
            Material = 1,
            None = 3 };

        public static class heading
        {
            public const int s_no = 0, ref_no = 1, item_desc = 2, rev = 3,
          planned = 4, actual = 5, return_date = 6, action = 7,
                comments = 8;
        }

    }
}
