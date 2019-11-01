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
    public class MixTimeSheetHandlerTests
    {

        private static Boolean compare
         (String mepStyle_employeeNo, String multiTrans_employeeNo)
        {
            //String mep_Trimmed = mepStyle_employeeNo.Trim('/');
            String mep_Trimmed = mepStyle_employeeNo
                .Substring(mepStyle_employeeNo.IndexOf('/') + 1);

             mep_Trimmed = mep_Trimmed.Trim('0');
            //eg
            //before = 02/04532
            //after trim = 4532

            mep_Trimmed = mep_Trimmed.Trim(); //trim unwanted white space

            String multiTrans_trimmed = multiTrans_employeeNo.Trim();  //remove white spaces

            multiTrans_trimmed = multiTrans_trimmed.TrimStart(new char[] { '0' });

            if (mep_Trimmed == multiTrans_trimmed)
                return true;
            else
                return false;

        }

        [TestMethod()]
        public void compare_multiTrans_employeeNo_to_MepStyle_employeeNo()
        {

            String assume_mep_emp_no = "2/15588";
            String assume_multi_emp_no = "000015588";

           if( compare(assume_mep_emp_no, assume_multi_emp_no)
                ==true)
            {
                Boolean is_true = true;
            }
           else
            {
                Boolean is_true = false;
            }

            //MixTimeSheetHandler mixTimeSheetHandler = new MixTimeSheetHandler();
            //PrivateObject obj = new PrivateObject(mixTimeSheetHandler);
            //var retVal = obj.Invoke("PrivateMethod");
            //Assert.AreEqual(expectedVal, retVal);

            Assert.Fail();
        }
        [TestMethod()]
        public void Add_Missing_data_from_mepStyle_to_MultiTransTest()
        {
            string assumed_checkin = "4:00";
            string assumed_overtime = "3";
            DateTime assumed_checkintime;
          DateTime.TryParse(assumed_checkin, out assumed_checkintime);

            MixTimeSheetHandler mixTimeSheetHandler = new MixTimeSheetHandler();
            PrivateObject obj = new PrivateObject(mixTimeSheetHandler);
            var retVal = obj.Invoke("calculate_checkOut_time",
                assumed_checkintime,assumed_overtime);

            
            Assert.Fail();
        }

        [TestMethod()]
        public void Add_siteNo_from_DailyTrans_to_MultiTransTest()
        {
            Assert.Fail();
        }
    }
}