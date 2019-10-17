using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander
{
  public  class HeadingWrap
    {

        public Excel.Range fullCell;
        public String headingName;

        public HeadingWrap(string headingName)
        {
            this.headingName = headingName;
        }
    }
}
