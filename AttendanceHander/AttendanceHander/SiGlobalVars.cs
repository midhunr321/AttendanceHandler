﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using AttendanceHander.MultipleTransaction;
using AttendanceHander.DailyTransactions;

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
        public Excel.Worksheet mepStyleCurrentWorkSheet;
        public List< MepStyleWrap> mepStyleWraps;
        public MepStyleHelper.Headings mepStyleHeadings;
        public DateTime mepStyleTimesheetMonthYear;

    
        public Excel.Workbook multiTransWorkbook;
        public Excel.Worksheet multiTransCurrentWorkSheet;
        public List<MultiTransWrap> multiTransWraps;
        public MultipleTransaction.MultiTransHelper.MultiHeadings multiTransHeadings;


        public Excel.Workbook dailyTransWorkbook;
        public Excel.Worksheet dailyTransCurrentWorkSheet;
        public List<DailyTransWrap> dailyTransWrap;
        public DailyTransHelper.Headings dailyTransHeadings;

    }


}
