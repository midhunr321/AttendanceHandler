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

        [TestMethod()]

        private Excel.Workbook openPlumbersTimesheet()
        {
            FormMain formMain = new FormMain();
            PrivateObject privateObject = new PrivateObject(formMain);

            return null;
        }
        public void MAIN_understand_the_excel_sheetTest()
        {
            //SET ==================================

            


            //ACT=======================================

            //ASSERT===================================
            Assert.Fail();
        }
    }
}