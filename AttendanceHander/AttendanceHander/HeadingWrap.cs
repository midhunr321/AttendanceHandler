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

        public override bool Equals(object obj)
        {
            if(obj is HeadingWrap)
            {
                var thatObj = obj as HeadingWrap;

                //when both fullcells are null, equals is true
                //but only if one fullcell is null, then equals is false

                if (this.fullCell == null && thatObj.fullCell == null &&
                    this.headingName == thatObj.headingName)
                    return true;
                else
                {
                    if (this.fullCell == null || thatObj.fullCell == null)
                        return false;

                    if (this.fullCell.Equals(thatObj.fullCell) &&
                        this.headingName == thatObj.headingName)
                        return true;
                }

            }
            return false;
        }
    }
}
