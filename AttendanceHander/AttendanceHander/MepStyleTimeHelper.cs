using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    class MepStyleTimeHelper
    {
        public LongItemWrap serialNo;
        public LongItemWrap code; //Employee no.
        public StrItemWrap name;
        public StrItemWrap designation;
        public StrItemWrap siteNo;
        public LongItemWrap totalOvertime;
        public List<DateOvertime> dateOvertimes;

        public class Heading : IEnumerable<String>
        {
            public  String mepStyleHeading = "Plumbers - Time Sheet";
            public String serialNo = "S. No";
            public String code = "Code";
            public String name = "Name";
            public String designation = "Design";
            public String siteNO = "Site Nos.";

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

        private Boolean check_the_loaded_excel_is_MEP_style_time()
        {

        }
     
        private Boolean find_the_heading_cell()
        {

        }


    }
}
