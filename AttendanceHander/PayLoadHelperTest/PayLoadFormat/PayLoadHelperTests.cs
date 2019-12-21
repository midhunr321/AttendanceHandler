using Microsoft.VisualStudio.TestTools.UnitTesting;
using AttendanceHander.PayLoadFormat;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander.PayLoadFormat.Tests
{
    [TestClass()]
    public class PayLoadHelperTests
    {
        [TestMethod()]
        public void PayLoadHelperTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void MAIN_understand_the_excel_sheetTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void calculateWorkTime()
        {
            TimeSpanItemWrap assumedTimeSpan = new TimeSpanItemWrap();
            assumedTimeSpan.content = new TimeSpan(8, 0, 0);

            MixTimeSheetHandler.WorkTimeCalculatedWarp result =
                new MixTimeSheetHandler.WorkTimeCalculatedWarp();
           
          result =   PayLoadHelper.Calculate_worktime_from_bioTotalWorkTime(assumedTimeSpan, true);

            Assert.Fail();
        }
    }
}