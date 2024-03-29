﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using AttendanceHander;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using AttendanceHander.DailyTransactions;
using AttendanceHander.MultipleTransaction;


namespace AttendanceHander.Tests
{
    [TestClass()]
    public class MepStyleSiteNoCodeAnalyzerTests
    {

        [TestMethod()]
        public void Test_ref_objects()
        {
            List<MultiTransWrap> assumed_multiTrans = new List<MultiTransWrap>();

            MultiTransWrap multiTransWrap1 = new MultiTransWrap();
            multiTransWrap1.firstName.content = "first name";
            MultiTransWrap multiTransWrap2 = new MultiTransWrap();
            multiTransWrap2.firstName.content = "second name";


        }

        [TestMethod()]
        public void DatetimeParse()
        {

            String assumeDateString = "15/12/2019";
            Nullable<DateTime> assumedDate = DateTime.Parse(assumeDateString);

            String assumed_string
                         = "17:05";
            DateTime extractedTime;
            DateTime.TryParse(assumed_string, out extractedTime);

            //if (DateTime.TryParseExact(assumed_string,
            //"HH:mm", CultureInfo.InvariantCulture,
            //System.Globalization.DateTimeStyles.AdjustToUniversal,
            //out extractedDAte)

            DateTime result =
                DateTimeHandler.mix_different_date_and_time((DateTime)assumedDate, extractedTime);

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