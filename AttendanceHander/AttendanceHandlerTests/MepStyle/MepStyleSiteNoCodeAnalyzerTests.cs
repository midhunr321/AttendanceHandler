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

            Assert.Fail();
        }

        [TestMethod()]
        public void analyze_stringTest()
        {
            //Arrange
            String code = "11to12_M273";

            //ACT
            DateTime monthyear = new DateTime(2019, 10, 11);
            MepStyleSiteNoCodeAnalyzer mepStyleSiteNoCodeAnalyzer =
                new MepStyleSiteNoCodeAnalyzer(monthyear);

            //Assert
            MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap expectedREsult
                = new MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap();
            expectedREsult.siteNo = "M273";
            expectedREsult.transferEndDate = new DateTime(2019, 10, 12);
            expectedREsult.transferStartDate = new DateTime(2019, 10, 11);

            var actualResult = mepStyleSiteNoCodeAnalyzer.analyze_string(code); 
           

            Assert.Equals(expectedREsult, actualResult);
        }

        [TestMethod()]
        public void invalidate_siteNoTest()
        {
            Assert.Fail();
        }
    }
}