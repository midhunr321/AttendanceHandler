using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using AttendanceHander.DailyTransactions;
using AttendanceHander.MultipleTransaction;
using System.Windows.Forms;
using System.Diagnostics;

namespace AttendanceHander
{
   public class MixTimeSheetHandler
    {
        class IntStringHolder
        {
            public int hour_inInt;
            public String time_inString;

        }

        private Boolean add_data_from_mep_to_multiTrans(MepStyleWrap mepStyleWrap,
            List<MultiTransWrap> multiTransWraps)
        {

            foreach (var multiWrap in multiTransWraps)
            {
                if (StringHandler.trim_and_compare_strings(mepStyleWrap.name.content,
                   multiWrap.firstName.content) &&
                   CommonOperations
                   .compare_multiTrans_employeeNo_to_MepStyle_employeeNo
                   (mepStyleWrap.code.content,
                   multiWrap.personnelNo.content)
                   )
                {
                    //that is employee no and ame is similar 
                    IntStringHolder mostRepeatedCheckInTime = new IntStringHolder();
                    mostRepeatedCheckInTime
               = get_mostRepeated_checkIn_time(multiWrap.personnelNo.content, multiTransWraps);


                    foreach (var dateOvertime in mepStyleWrap.dateOvertimes)
                    {
                        //now we need to compare the overtime dates

                        if (DateTimeHandler
                             .Compare_dates_only(dateOvertime.date, multiWrap.date.content.Value))
                        {
                            if (convert_and_assign_mep_style_overtime_to_biometric_style_overtime
                                  (dateOvertime, multiWrap.date,
                                  mostRepeatedCheckInTime.time_inString, multiWrap)
                                  == false)
                                return false;
                            
                        }
                    }

                }
            }

            return true;
        }

        private IntStringHolder get_mostRepeated_checkIn_time(String employeeNo,
            List<MultiTransWrap> multiTransWraps)
        {
            List<MultiTransWrap> fullDataOfEmployee
                = multiTransWraps.Where(x => x.personnelNo.content == employeeNo
                ).ToList();


            int mostRepeatedCheckInHour = (from item in fullDataOfEmployee
                                           group item by item.checkInTime1.content.Value.Hour into group_by_checkin_hr
                                           orderby group_by_checkin_hr.Count() descending
                                           select group_by_checkin_hr.Key).First();

            IntStringHolder intStringHolder = new IntStringHolder();
            intStringHolder.hour_inInt = mostRepeatedCheckInHour;
            String checkInTime = mostRepeatedCheckInHour.ToString()
                + ":00";
            intStringHolder.time_inString = checkInTime;

            return intStringHolder;

        }

        private Nullable<DateTime> calculate_checkOut_time(DateTime checkIn_time, String overtime)
        {
            Nullable<DateTime> checkOut_time=null;

            int normalWorking_hours = SiGlobalVars.Instance.assumed_normal_workingHours;

            StringHandler stringHandler = new StringHandler();
           
            if(stringHandler
                .is_this_string_alpha_numeric_or_numeric_or_alpha_only(overtime)
                == All_const.str_type.Numeric)
            {
                int overtime_Int
                  = int.Parse(overtime);

                checkOut_time = checkIn_time.AddHours(overtime_Int);
                //sometimes overtime will be like 8 or 10 or 5 etc
               
            }
            else if(overtime == SiGlobalVars.Instance.assumed_SickLeave_key)
            {
                //for sick leave...like "SL"
                checkOut_time = checkIn_time.AddHours(normalWorking_hours);
                // for sickleave, no overtime is provided;
            }

            return checkOut_time;
        }
        private Boolean convert_and_assign_mep_style_overtime_to_biometric_style_overtime
            (DateOvertime mepOvertime, DateItemWrap multiTransOvertime,
            String mostRepeated_checkIn_time, MultiTransWrap multiTransWrap)
        {
            DateTime checkIn_time;
            
            if(DateTime.TryParse(mostRepeated_checkIn_time,out checkIn_time)
                == false)
            {
                MessageBox.Show("Couldn't convert Most Repeated CheckIn-time = " +
                    mostRepeated_checkIn_time);
                StackTrace stack = new StackTrace();
                Console.WriteLine(stack.ToString());
                return false;
            }
            else
            {

             var checkOut_time=
                    calculate_checkOut_time(checkIn_time,mepOvertime.overtime);

                if (checkOut_time == null)
                {
                    MessageBox.Show("Couldn't calculate checkout time ");
                    StackTrace stackTrace = new StackTrace();
                    Console.WriteLine(stackTrace.ToString());
                    return false;
                }

                checkIn_time = DateTimeHandler
                      .mix_different_date_and_time(multiTransWrap.date.content.Value,
                      checkIn_time);

                checkOut_time = DateTimeHandler
                    .mix_different_date_and_time(multiTransWrap.date.content.Value, checkOut_time.Value);

                multiTransWrap.checkInTime1.content = checkIn_time;
                multiTransWrap.checkInTime1.contentInString = checkIn_time.TimeOfDay.ToString();

                multiTransWrap.checkOutTime2.content = checkOut_time.Value;
                multiTransWrap.checkOutTime2.contentInString = checkOut_time.Value.TimeOfDay.ToString();

                TimeSpan timeSpan = checkOut_time.Value.Subtract(checkIn_time);
                multiTransWrap.workingTime1.content = timeSpan;
                multiTransWrap.workingTime1.contentInString = timeSpan.ToString();

                return true;

            }

        }

        public Boolean Add_Missing_data_from_mepStyle_to_MultiTrans(List<MepStyleWrap> mepStyleWraps,
           ref List<MultiTransWrap> multiTransWraps)
        {
            foreach (var mepWrap in mepStyleWraps)
            {
                if (add_data_from_mep_to_multiTrans(mepWrap, multiTransWraps)
                    == false)
                    return false;
            }

            return true;

        }

        private Boolean add_site_no_to_multiTrans_sheet(MultiTransWrap multiTransWrap)
        {
            if (MultiTransHelper
                   .Add_a_heading_column_for_site_no
                   (SiGlobalVars.Instance.multiTransHeadings.totalTimeWorked.fullCell,
                   SiGlobalVars.Instance.multiTransCurrentWorkSheet,
                  ref SiGlobalVars.Instance.multiTransHeadings)
                   == true)
            {
                //true means already site no heading available
                if (multiTransWrap.siteNo == null)
                    multiTransWrap.siteNo = new StrItemWrap();

                //for each row..the site no data cell will always be after totalTimeWorked cell.
                multiTransWrap.siteNo.fullCell = multiTransWrap.totalTimeWorked.fullCell.Next;
                //now assign site no value
                multiTransWrap.siteNo.fullCell.Value = multiTransWrap.siteNo.content;
                return true;
            }
            else
            {
                MessageBox.Show("Couldn't add site no for the plumber Name-cell= "
                    + multiTransWrap.firstName.fullCell.Address.ToString());
                return false;
            }

        }
        private Boolean assign_data_to_multiTransWrap_and_add_siteNo_to_MultiTrans
      (List<DailyTransWrap> dailyTransWraps,
       MultiTransWrap multiWrap)
        {

            foreach (var dailyWrap in dailyTransWraps)
            {
                if (StringHandler.trim_and_compare_strings(dailyWrap.firstName.content,
                    multiWrap.firstName.content) &&
                    StringHandler.trim_and_compare_strings(dailyWrap.date.contentInString,
                    multiWrap.date.contentInString) &&
                     StringHandler.trim_and_compare_strings(dailyWrap.personnelNo.content,
                    multiWrap.personnelNo.content)
                    )
                {
                    if (multiWrap.siteNo == null)
                        multiWrap.siteNo = new StrItemWrap();

                    multiWrap.siteNo.content = dailyWrap.deviceName.content;
                    if (add_site_no_to_multiTrans_sheet(multiWrap)
                         == false)
                        return false;

                    return true;
                }


            }
            return false;
        }
        public static Boolean Add_siteNo_from_DailyTrans_to_MultiTrans
            (List<DailyTransWrap> dailyTransWraps,
           ref List<MultiTransWrap> multiTransWraps)
        {
            MixTimeSheetHandler mixTimeSheetHandler = new MixTimeSheetHandler();
            foreach (var multiwrap in multiTransWraps)
            {
                mixTimeSheetHandler.assign_data_to_multiTransWrap_and_add_siteNo_to_MultiTrans
                     (dailyTransWraps, multiwrap);
            }

            return true;

        }
    }
}
