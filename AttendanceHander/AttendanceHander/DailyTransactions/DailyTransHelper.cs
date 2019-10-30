using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander.DailyTransactions
{
   public class DailyTransHelper
    {
        private Excel.Worksheet worksheet;
        private Excel.Workbook workbook;

        public DailyTransHelper(Excel.Worksheet worksheet, Excel.Workbook workbook)
        {
            this.worksheet = worksheet;
            this.workbook = workbook;
        }

        public class Headings : IEnumerable<HeadingWrap>
        {
            public HeadingWrap sheetHeading =
                new HeadingWrap("Transactions");
            public HeadingWrap personnelNo =
                new HeadingWrap("Personnel No.");
            public HeadingWrap firstName = new HeadingWrap("First Name");
            public HeadingWrap lastName = new HeadingWrap("Last Name");
            public HeadingWrap position = new HeadingWrap("Position");
            public HeadingWrap department = new HeadingWrap("Department");
            public HeadingWrap date = new HeadingWrap("Date");
            public HeadingWrap time = new HeadingWrap("Time");
            public HeadingWrap punchStatus = new HeadingWrap("Punch Status");
            public HeadingWrap workCode = new HeadingWrap("Work Code");
            public HeadingWrap gpsLocation = new HeadingWrap("GPS Location");
            public HeadingWrap area = new HeadingWrap("Area");
            public HeadingWrap deviceName = new HeadingWrap("Device Name");
            public HeadingWrap deviceSerialNo = new HeadingWrap("Device SN");
            public HeadingWrap dataFrom = new HeadingWrap("Data From");

            public IEnumerator<HeadingWrap> GetEnumerator()
            {

                return (new List<HeadingWrap>()
                {sheetHeading,personnelNo,
                    firstName,lastName,position,department,
                    date,time,punchStatus,workCode, gpsLocation,
                area,deviceName, deviceSerialNo,dataFrom  }.GetEnumerator());
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }


        }




    }
}
