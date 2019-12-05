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
    public class CommonOperationsTests
    {
        [TestMethod()]
        public void modify_value_in_cellTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void convert_siteNo_to_SiteNoMechFormatTest()
        {
            CommonOperations commonOperations = new CommonOperations(null);
            StrItemWrap mock_deviceName;
            mock_deviceName = new StrItemWrap();
            mock_deviceName.content = "S276-1101";
          String result=  commonOperations.convert_siteNo_to_SiteNoMechFormat(mock_deviceName);

            Assert.Fail();
        }

        [TestMethod()]
        public void compare_multiTrans_employeeNo_to_MepStyle_employeeNoTest()
        {
        }

        [TestMethod()]
        public void CommonOperationsTest()
        {
        }

        [TestMethod()]
        public void filter_searchResult_by_comparing_row_no_of_adjacent_headingsTest()
        {
        }

        [TestMethod()]
        public void feed_time_data_to_dataWrapTest()
        {
        }

        [TestMethod()]
        public void feed_time_data_to_dataWrapTest1()
        {
        }

        [TestMethod()]
        public void employeeNo_is_validTest()
        {
        }

        [TestMethod()]
        public void name_is_validTest()
        {
        }
    }
}