﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander.PayLoadFormat
{
   public class PayLoadWrap
    {
        public List<Day> days;
        public class Day
        {

            public Excel.Worksheet sheet;
            public List<Employee> employees;
            public StrItemWrap company;
            public DateItemWrap date;
            public StrItemWrap section;
            public StrItemWrap job;

            public Day(Excel.Worksheet sheet)
            {
                this.sheet = sheet;
            }

            public class Employee
            {
               

                public StrItemWrap serialNo;
                public StrItemWrap code;
                public StrItemWrap name;
                public StrItemWrap design;
                public StrItemWrap job_siteNo;
                public DecimalItemWrap workTime;
                public DecimalItemWrap noBreak;
                public DecimalItemWrap overTime;
            }
  
           

        }
    }
}
