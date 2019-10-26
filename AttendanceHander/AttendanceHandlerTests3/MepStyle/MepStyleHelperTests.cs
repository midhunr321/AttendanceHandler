using Microsoft.VisualStudio.TestTools.UnitTesting;
using AttendanceHander;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace AttendanceHander.Tests
{
    [TestClass()]
    public class MepStyleHelperTests
    {
        public void MepStyleHelperTest()
        {

            Assert.Fail();
        }


        private Excel.Workbook openPlumbersTimesheet()
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks
                .Open("C:\\Users\\midhun\\Source\\Repos\\AttendanceHandler\\AttendanceHander" +
                "\\TESTING - Plumbers 2019.xlsx");


            return workbook;
        }

        [TestMethod()]
        public void MAIN_understand_the_excel_sheetTest()
        {
            //SET ==================================
            Excel.Workbook workbook;
            workbook = openPlumbersTimesheet();
            int lastSheet = workbook.Sheets.Count - 1;
            Excel.Worksheet worksheet = workbook.Sheets[lastSheet];

            SiGlobalVars.Instance.mepStyleWorkbook = workbook;
            SiGlobalVars.Instance.mepStyleCurrentMonthWorkSheet = worksheet;

            MepStyleHelper mepStyleHelper = new MepStyleHelper(workbook, worksheet);
            mepStyleHelper.MAIN_understand_the_excel_sheet();
            //ACT=======================================

            //ASSERT===================================
            Assert.Fail();
        }
    }
}