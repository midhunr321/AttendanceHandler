using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander
{
    public sealed class SiGlobalVars
    {
        private static volatile SiGlobalVars instance;
        private static object syncRoot = new Object();

        private SiGlobalVars() { }

        public static SiGlobalVars Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new SiGlobalVars();
                    }
                }

                return instance;
            }
        }

        public Excel.Workbook mepStyleWorkbook;
        public Excel.Worksheet mepStyleCurrentMonthWorkSheet;
        public List< MepStyleWrap> mepStyleWraps;
        public MepStyleHelper.Headings mepStyleHeadings;
    }


}
