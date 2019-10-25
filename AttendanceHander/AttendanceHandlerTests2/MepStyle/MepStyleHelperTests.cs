using Microsoft.VisualStudio.TestTools.UnitTesting;
using AttendanceHander;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander.Tests
{
    [TestClass()]
    public class MepStyleHelperTests
    {
        [TestMethod()]
        public void MepStyleHelperTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void MAIN_understand_the_excel_sheetTest()
        {
            //INITIALIZE FOR TESTING================================
            FormMain formMain = new FormMain();
            PrivateObject privateObject = new PrivateObject(formMain);
            privateObject.Invoke("openFile", true);



            //ACT============================================

            //ASSERT===============================================

            Assert.Fail();
        }

    
    }
}