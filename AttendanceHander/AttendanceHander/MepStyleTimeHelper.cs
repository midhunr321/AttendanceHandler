using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        private MepStyleTimeHelper.Headings headings;
        private Excel.Worksheet worksheet;
        private Excel.Workbook workbook;

        public MepStyleTimeHelper(MepStyleTimeHelper.Headings headings,
            Excel.Workbook workbook, Excel.Worksheet worksheet)
        {
            this.headings = headings;
            this.workbook = workbook;
            this.worksheet = worksheet;
        }

        public class Headings : IEnumerable<HeadingWrap>
        {
            public HeadingWrap mepStyleHeading =
                new HeadingWrap("Plumbers - Time Sheet");
            public HeadingWrap serialNo = new HeadingWrap("S. No");
            public HeadingWrap code = new HeadingWrap("Code");
            public HeadingWrap name = new HeadingWrap("Name");
            public HeadingWrap designation = new HeadingWrap("Design");
            public HeadingWrap siteNO = new HeadingWrap("Site Nos.");

            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {mepStyleHeading,serialNo,
                    code,name,designation,siteNO }.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
        public void understand_the_excel_sheet()
        {
            if (headings == null)
            {
                MessageBox.Show("Heading Strings is null");
                return;
            }

            find_the_heading_cell(headings);
        }
        private Boolean check_the_loaded_excel_is_MEP_style_time()
        {

        }

        private Boolean find_the_heading_cells(MepStyleTimeHelper.Headings headings)
        {
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);
            foreach (HeadingWrap heading in headings)
            {
                List<Excel.Range> temp_heading = new List<Excel.Range>();
                temp_heading = 
                    eXCEL_HELPER.find_fix_column_heading(heading.headingName,
                    Excel.XlSearchDirection.xlNext,
                    Excel.XlSearchOrder.xlByRows);
                //TODO: Should carryout the search from the top

                // if the search count is not more than 1 then,
                if (temp_heading != null && temp_heading.Count == 1)
                {
                    if (temp_heading[0] != null)
                        heading.fullCell = temp_heading[0];
                }
                else
                {
                    //TODO: if more than one search results
                    //we need to filter it out
                    //like check if the full cell is within the same heading row
                    //that way we can filter out other results.

                }

            }
        }


    }
}
