using Microsoft.VisualStudio.TestTools.UnitTesting;
using AttendanceHander;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace AttendanceHander.Tests
{
    [TestClass()]
    public class MepStyleSiteNoCodeAnalyzerTests
    {
        [TestMethod()]
        public void DatetimeParse()
        {

            String assumed_string
                         = "12:12";
            DateTime extractedDAte;
            if (DateTime.TryParseExact(assumed_string,
                "HH:mm", CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AdjustToUniversal,
                out extractedDAte)
            == false)
                Assert.Fail();
        }
        [TestMethod()]
        public void MepStyleSiteNoCodeAnalyzerTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void analyze_stringTest()
        {
            //Arrange
            String code = "28to31_M274";

            //ACT
            DateTime monthyear = new DateTime(2019, 10, 11);
            MepStyleSiteNoCodeAnalyzer mepStyleSiteNoCodeAnalyzer =
                new MepStyleSiteNoCodeAnalyzer(monthyear);

            //Assert
            MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap expectedREsult
                = new MepStyleSiteNoCodeAnalyzer.ExtractedDataWrap();
            expectedREsult.siteNo = "M274";
            expectedREsult.transferStartDate = new DateTime(2019, 10, 28);
            expectedREsult.transferEndDate = new DateTime(2019, 10, 31);
           

            var actualResult = mepStyleSiteNoCodeAnalyzer.analyze_string(code);

            Assert.AreEqual(expectedREsult, actualResult);

        }

        [TestMethod()]
        public void invalidate_siteNoTest()
        {
            Assert.Fail();
        }
    }
}