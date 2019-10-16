using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
    class MultipleTransHelper
    {
        public class Heading: IEnumerable<String>
        {
            public static String SERIAL_NO = "NO";

            public IEnumerator<String> GetEnumerator()
            {

                return (new List<String>()
                {SERIAL_NO }.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

        }

        private void check_if_current_sheet_is_multiple_transaction()
        {

        }
    }
}
