using Microsoft.VisualStudio.TestTools.UnitTesting;
using AttendanceHander;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander.Tests
{
    [TestClass()]
    public class MepStyleSiteNoCodeAnalyzerTests
    {
        [TestMethod()]
        public void MepStyleSiteNoCodeAnalyzerTest()
        {
            //Arrange
            int expected_startDate = 11;
            int expected_endDate = 12;
            String code = "11to12_M273";
            String site = "M273";
            DateTime timesheetMonth = new DateTime(11, 10, 2019);
            //ACT
            MepStyleSiteNoCodeAnalyzer mepStyleSiteNoCodeAnalyzer
                = new MepStyleSiteNoCodeAnalyzer(timesheetMonth);
            mepStyleSiteNoCodeAnalyzer.analyze_string(code);

            //Assert
            MepStyleSiteNoCodeAnalyzer mep = new MepStyleSiteNoCodeAnalyzer(new DateTime());
            

            Assert.Fail();
        }

        [TestMethod()]
        public void analyze_stringTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void invalidate_transfer_dateTest()
        {
            Assert.Fail();
        }
    }
}