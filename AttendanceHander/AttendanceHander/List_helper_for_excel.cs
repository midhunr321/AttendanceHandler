using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;

namespace AttendanceHander
{
    class List_helper_for_excel
    {

        public Excel.Range find_lowest_number_in_list(List<Excel.Range> list)
        {
            Excel.Range lowest =(Excel.Range) list[0];
            //first assume very first element is the lowest the list array.
            foreach (Excel.Range range in list)
            {

                if (range.Column < lowest.Column && range.Row < lowest.Row)
                    lowest = range;
            }
            return lowest;
        }
    }
}
