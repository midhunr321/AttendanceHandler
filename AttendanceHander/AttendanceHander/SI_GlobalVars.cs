using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander
{
    public sealed class SI_GlobalVars
    {
        private static volatile SI_GlobalVars instance;
        private static object syncRoot = new Object();

        private SI_GlobalVars() { }

        public static SI_GlobalVars Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new SI_GlobalVars();
                    }
                }

                return instance;
            }
        }

        public Excel.Workbook mepStyleTimeSheet;
        public Excel.Worksheet mepStyleCurrentMonth;
    }


}
