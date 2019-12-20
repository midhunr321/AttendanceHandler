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
using AttendanceHander.PayLoadFormat;

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
            List<String> skipList_Emp_no = new List<String>();

            foreach (var multiWrap in multiTransWraps)
            {
                if (skipList_Emp_no.Contains(multiWrap.personnelNo.content))
                    continue; //if skip list contains this employee no, then skip

                if (StringHandler.trim_and_compare_strings(mepStyleWrap.name.content,
                   multiWrap.firstName.content) &&
                   CommonOperations
                   .compare_multiTrans_employeeNo_to_MepAndPayLoad_employeeNo
                   (mepStyleWrap.code.content,
                   multiWrap.personnelNo.content)
                   )
                {
                    //that is employee no and ame is similar 
                    IntStringHolder mostRepeatedCheckInTime = new IntStringHolder();
                    mostRepeatedCheckInTime
               = get_mostRepeated_checkIn_time(multiWrap.personnelNo.content, multiTransWraps,
               multiWrap);
                    if (mostRepeatedCheckInTime == null)
                    {
                        skipList_Emp_no.Add(multiWrap.personnelNo.content);
                        //if get_mostRepeated_checkIn_time() returns null
                        //it means not even single check-in time is available with this guy
                        //in this case we need to skip this employee completely.
                        continue;

                    }

                    foreach (var dateOvertime in mepStyleWrap.dateOvertimes)
                    {
                        //now we need to compare the overtime dates
                        if (is_overTime_in_mep_style_valid_or_nonEmpty(dateOvertime)
                              == false)
                            continue; //if overtime is empty in mep style then continue the iteration

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

        private bool is_overTime_in_mep_style_valid_or_nonEmpty(DateOvertime dateOvertime)
        {

            if (String.IsNullOrEmpty(dateOvertime.overtime))
                return false;
            else if (String.IsNullOrWhiteSpace(dateOvertime.overtime))
                return false;

            StringHandler stringHandler = new StringHandler();
            if (stringHandler
                .is_this_string_alpha_numeric_or_numeric_or_alpha_only
                (dateOvertime.overtime) == All_const.str_type.Numeric)
                return true;
            else if (dateOvertime.overtime == SiGlobalVars.Instance.assumed_SickLeave_key.Key)
                return true;


            return false;
        }

        private IntStringHolder get_mostRepeated_checkIn_time(String employeeNo,
            List<MultiTransWrap> multiTransWraps, MultiTransWrap multiTransWrap)
        {
            List<MultiTransWrap> fullDataOfEmployee
                = multiTransWraps.Where(x => x.personnelNo.content == employeeNo
                && x.checkInTime1.content != null).ToList();
            //we don't want those datas where checkIntime is null.
            if (fullDataOfEmployee.Count == 0)
            {
                MessageBox.Show("No default or routing CheckIn time was found for employee = " +
                    employeeNo + " for the date = " + multiTransWrap.date.contentInString
                    + "Thus it is skipped");
                return null;
            }


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
            Nullable<DateTime> checkOut_time = null;

            int normalWorking_hours = SiGlobalVars.Instance.assumed_normal_workingHours;

            StringHandler stringHandler = new StringHandler();

            if (stringHandler
                .is_this_string_alpha_numeric_or_numeric_or_alpha_only(overtime)
                == All_const.str_type.Numeric)
            {
                int overtime_Int
                  = int.Parse(overtime);
                //first add normal working time
                checkOut_time = checkIn_time.AddHours(normalWorking_hours);
                //then add overtime
                checkOut_time = checkOut_time.Value.AddHours(overtime_Int);
                //sometimes overtime will be like 8 or 10 or 5 etc

            }


            return checkOut_time;
        }
        private Boolean convert_and_assign_mep_style_overtime_to_biometric_style_overtime
            (DateOvertime mepOvertime, DateItemWrap multiTransOvertime,
            String mostRepeated_checkIn_time, MultiTransWrap multiTransWrap)
        {
            //check for sick leave
            if (mepOvertime.overtime == SiGlobalVars.Instance.assumed_SickLeave_key.Key)
            {
                //means sick leave
                //in this case we can skip the check in check out part
                //we can write sick leave in the total work time
                CommonOperations.modify_value_in_cell(multiTransWrap.totalTimeWorked
                  .fullCell, SiGlobalVars.Instance.assumed_SickLeave_key.Value,
                  SiGlobalVars.Instance.assumed_editFont_colour);
                return true;

            }

            DateTime checkIn_time;


            if (DateTime.TryParse(mostRepeated_checkIn_time, out checkIn_time)
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

                var checkOut_time =
                       calculate_checkOut_time(checkIn_time, mepOvertime.overtime);

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
                CommonOperations.modify_value_in_cell(multiTransWrap.checkInTime1
                    .fullCell, multiTransWrap.checkInTime1.contentInString,
                    SiGlobalVars.Instance.assumed_editFont_colour);

                multiTransWrap.checkOutTime1.content = checkOut_time.Value;
                multiTransWrap.checkOutTime1.contentInString = checkOut_time.Value.TimeOfDay.ToString();
                CommonOperations.modify_value_in_cell(multiTransWrap.checkOutTime1
                   .fullCell, multiTransWrap.checkOutTime1.contentInString,
                   SiGlobalVars.Instance.assumed_editFont_colour);

                TimeSpan timeSpan = checkOut_time.Value.Subtract(checkIn_time);
                timeSpan = DateTimeHandler.get_absolute_timeSpan(timeSpan);

                multiTransWrap.workingTime1.content = timeSpan;
                multiTransWrap.workingTime1.contentInString = timeSpan.ToString();
                CommonOperations.modify_value_in_cell(multiTransWrap.workingTime1
                   .fullCell, multiTransWrap.workingTime1.contentInString,
                   SiGlobalVars.Instance.assumed_editFont_colour);


                multiTransWrap.totalTimeWorked.content = timeSpan;
                multiTransWrap.totalTimeWorked.contentInString = timeSpan.ToString();
                CommonOperations.modify_value_in_cell(multiTransWrap.totalTimeWorked
                   .fullCell, multiTransWrap.totalTimeWorked.contentInString,
                   SiGlobalVars.Instance.assumed_editFont_colour);

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

        private Boolean add_site_no_to_multiTrans_sheet(
            MultiTransWrap multiTransWrap)
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
                    multiTransWrap.siteNo = new MultiTransWrap.StrItemWrap();

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
                        multiWrap.siteNo = new MultiTransWrap.StrItemWrap();

                    multiWrap.siteNo.content = dailyWrap.deviceName.content;
                    if (multiWrap.siteNoMechFormat == null)
                        multiWrap.siteNoMechFormat = new MultiTransWrap.SiteNoMechFormat();
                    if (multiWrap.siteNoMechFormat.longSiteNo == null)
                        multiWrap.siteNoMechFormat.longSiteNo = new MultiTransWrap.StrItemWrap();
                    if (multiWrap.siteNoMechFormat.shortName == null)
                        multiWrap.siteNoMechFormat.shortName = new MultiTransWrap.StrItemWrap();

                    //now we need to feed the site no in mech format
                    //that means we need to replace 'S' in site no with 'M'
                    //eg S269 = M269
                    //shortName means = M276
                    //fullName means = M276-1101 
                    if (dailyWrap.deviceName.content != String.Empty)
                    {
                        multiWrap.siteNoMechFormat.shortName.content = CommonOperations
                    .convert_siteNo_to_SiteNoMechFormat_ShortName(dailyWrap.deviceName);
                        multiWrap.siteNoMechFormat.longSiteNo.content = CommonOperations
                            .replace_first_S_in_siteNo_with_M(dailyWrap.deviceName);
                    }


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

        internal static Boolean Transfer_data_from_multiTrans_to_mepStyle(
            List<MultiTransWrap> multiTransWraps, List<MepStyleWrap> mepStyleWraps)
        {
            foreach (var multiWrap in multiTransWraps)
            {
                foreach (var mepWrap in mepStyleWraps)
                {
                    if (CommonOperations
                        .compare_multiTrans_employeeNo_to_MepAndPayLoad_employeeNo
                        (mepWrap.code.content, multiWrap.personnelNo.content)
                        == true
                        )
                    {
                        foreach (var dateOvertime in mepWrap.dateOvertimes)
                        {
                            if (DateTimeHandler
                                .Compare_dates_only(dateOvertime.date, multiWrap.date.content.Value))
                            {
                                dateOvertime.overtime = multiWrap.totalTimeWorked.contentInString;
                                dateOvertime.fullCell.Value = multiWrap.totalTimeWorked.contentInString;
                                if (multiWrap.siteNo != null)
                                    dateOvertime.siteNo = multiWrap.siteNo.content;
                            }
                        }
                    }


                }
            }
            return true;
        }

        internal Boolean correct_siteNo_in_multiTrans_with_fullMechSiteNo(List<MultiTransWrap> multiTransWraps)
        {
            Boolean not_even_one_SiteNo_is_corrected = true;
            foreach (var item in multiTransWraps)
            {
                //first read or get shortnames
                //first get the shortname
                if (item.siteNo == null)
                    continue;   //sometimes siteno is not available then ignore the cell for correction


                String fullName =
                    CommonOperations.replace_first_S_in_siteNo_with_M(item.siteNo);

                if (fullName == null)
                {
                    MessageBox.Show("Failed to correct Site No. Process is Aborted; No changes have " +
                        "been made in the multiTrans Excel File");
                    return false;
                }

                item.siteNoMechFormat.longSiteNo.content
                    = fullName;

            }
            foreach (var item in multiTransWraps)
            {
                if (item.siteNo == null)
                    continue;   //sometimes siteno is not available then ignore the cell for correction
                //now write 
                item.siteNo.fullCell.Value = item.siteNoMechFormat.longSiteNo.content;
                item.siteNoMechFormat.longSiteNo.fullCell = item.siteNo.fullCell;
                not_even_one_SiteNo_is_corrected = false;

            }

            if (not_even_one_SiteNo_is_corrected == true)
                return false;


            return true;
        }


        internal static Boolean Transfer_data_from_multiTrans_to_payLoad
            (List<DateTime> holidays, bool printBio_inPayLoad)
        {

            if (all_employees_of_MultiTrans_is_available_with_PayLoad
               (SiGlobalVars.Instance.multiTransWraps,
               SiGlobalVars.Instance.payLoadWrap) == false)
                return false;


            foreach (var multiWrap in SiGlobalVars.Instance.multiTransWraps)
            {
                //first we have to check if the site no is available for a particular 
                //employee for a particular date

                if (multiWrap.siteNoMechFormat.shortName == null)
                {
                    MessageBox.Show("The Site No is null for the Employee No = "
                        + multiWrap.personnelNo
                           + "; for the Date = " + multiWrap.date.contentInString);
                    return false;
                }

                foreach (var payLoadWrapDay in SiGlobalVars.Instance.payLoadWrap.days)
                {
                    if (payLoadWrapDay.sheet.Name.Trim()
                        == multiWrap.date.content.Value.Day.ToString().Trim())
                    {
                        //that is the dates are equal

                        payLoadWrapDay.date.contentInString = multiWrap.date.contentInString;
                        payLoadWrapDay.date.content = multiWrap.date.content;

                        //now check if this date is a holiday or friday
                        Boolean holidayOrFriday
                            = this_date_is_a_holidayOrFriday(payLoadWrapDay.date.content.Value,
                            holidays);

                        //now check for the employee codes
                        foreach (var payLoadWrapDayEmpl in payLoadWrapDay.employees)
                        {

                            if (CommonOperations
                                .compare_multiTrans_employeeNo_to_MepAndPayLoad_employeeNo
                                (payLoadWrapDayEmpl.code.content, multiWrap.personnelNo.content)
                                == true)
                            {
                                write_data_to_payLoadFormat_from_MultiTrans
                                    (payLoadWrapDayEmpl, multiWrap,
                                    holidayOrFriday,printBio_inPayLoad);
                            }

                        }


                    }
                }




            }
            return true;
        }

        private static bool all_employees_of_MultiTrans_is_available_with_PayLoad
            (List<MultiTransWrap> multiTransWraps, PayLoadWrap payLoadWrap)
        {
            foreach (var payLoadWrapDay in payLoadWrap.days)
            {
                foreach (var multiEmp in multiTransWraps)
                {
                    Boolean employeeFound = false;
                    foreach (var payLoadWrapDayEmp in payLoadWrapDay.employees)
                    {
                        if (CommonOperations
                            .compare_multiTrans_employeeNo_to_MepAndPayLoad_employeeNo
                            (payLoadWrapDayEmp.code.content, multiEmp.personnelNo.content))
                        {
                            employeeFound = true;
                            break;
                        }
                    }

                    if (employeeFound == false)
                    {
                        MessageBox.Show("Employee Code = " + multiEmp.personnelNo.content
                            + " is not found in PayLoad but the same is available in Multiple Transactions");
                        return false;
                    }
                }
            }

            return true;
        }

        internal static bool Transfer_MEPdata_to_payLoad()
        {
            if (SiGlobalVars.Instance.Holidays == null)
            {
                DialogResult dialogResult =
                      MessageBox.Show("Selected Holidays is null. Continue?", "Warning",
                      buttons: MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                    return false;
            }
            if (SiGlobalVars.Instance.mepStyleWraps == null ||
               SiGlobalVars.Instance.payLoadWrap == null)
            {
                return false;
            }

            //first we need to check if the payload sheet has all the employee codes
            //mentioned in the mepformat

            if (all_employees_of_MEPformat_is_available_with_PayLoad
                (SiGlobalVars.Instance.mepStyleWraps,
                SiGlobalVars.Instance.payLoadWrap) == false)
                return false;

            foreach (var mepWrap in SiGlobalVars.Instance.mepStyleWraps)
            {
                foreach (var mepDateOverWrap in mepWrap.dateOvertimes)
                {
                    foreach (var payLoadDay in SiGlobalVars.Instance.payLoadWrap.days)
                    {
                       

                        foreach (var payLoadDayEmp in payLoadDay.employees)
                        {
                            if (DateTimeHandler
                                .Compare_dates_only(payLoadDay.date.content.Value, mepDateOverWrap.date))
                            {
                                //ie same dates.
                                if (mepWrap.code.content == payLoadDayEmp.code.content)
                                {
                                    Boolean isHoliday = this_date_is_a_holidayOrFriday(payLoadDay.date.content.Value,
                                          SiGlobalVars.Instance.Holidays);
                                    write_data_to_payLoadFormat_from_mepStyle
                                        (payLoadDayEmp, mepWrap, isHoliday);
                                }

                            }

                        }
                    }
                }

            }

            return true;
        }

        private static bool all_employees_of_MEPformat_is_available_with_PayLoad
            (List<MepStyleWrap> mepStyleWraps, PayLoadWrap payLoadWrap)
        {
            foreach (var payLoadWrapDay in payLoadWrap.days)
            {
                foreach (var mepEmpl in mepStyleWraps)
                {
                    Boolean employee_found = false;

                    foreach (var payLoadWrapDayEmp in payLoadWrapDay.employees)
                    {
                        if (mepEmpl.code.content == payLoadWrapDayEmp.code.content)
                        {
                            employee_found = true;
                            break;
                        }


                    }

                    //if employee is not found in the payload format
                    //then it is an error

                    if (employee_found == false)
                    {
                        MessageBox.Show("Employee Code = " + mepEmpl.code
                            + " Couldn't be found in PayLoad Format but it is available in Mep Style Timesheet");
                        return false;
                    }

                }


            }

            return true;
        }

        private static bool this_date_is_a_holidayOrFriday(DateTime thisDate, List<DateTime> explicit_holidays)
        {
            foreach (var day in SiGlobalVars.Instance.DEFAULT_HOLIDAYS)
            {
                //check if fridays ..or like that
                if (thisDate.DayOfWeek == day)
                    return true; //means normal off day like friday

            }

            //now for explicit holidays which was selected in the form

            foreach (var explictHoliday in explicit_holidays)
            {
                if (DateTimeHandler.Compare_dates_only(thisDate, explictHoliday)
                    == true)
                    return true; //means a holiday like national day
            }

            return false;


        }

        public class WorkTimeCalculatedWarp
        {
            public class Wrap
            {
                public decimal content;

            }

            public Wrap workTimeHours;
            public Wrap noBreak;
            public Wrap overTime;
        }

        private static MixTimeSheetHandler.WorkTimeCalculatedWarp calculate_workTime_from_MEPStyleOvertime
           (DateOvertime mepOvertime, bool thisDate_is_fridayOrHoliday, MepStyleWrap mepStyleWrap)
        {
            MixTimeSheetHandler.WorkTimeCalculatedWarp workTimeCalculated
                = new MixTimeSheetHandler.WorkTimeCalculatedWarp();

            StringHandler stringHandler = new StringHandler();
            int result_totalWorkTime;

            if (thisDate_is_fridayOrHoliday == true)
            {
                //i.e friday means...absents don't count

                //if there is some overtime then add it.
                int converted;
                if (int.TryParse(mepOvertime.overtime, out converted) == true)
                {
                    //if friday means all worktime will be converted to overtime
                    //so normal worktime will be zero
                    workTimeCalculated.overTime.content
                        = (int)SiGlobalVars.Instance.DEFAULT_WORKING_HOURS + converted;

                    workTimeCalculated.workTimeHours.content = 0;
                }
                else
                {
                    //means friday but the employee didn't work. so
                    //overtime will be normal 8 hours
                    workTimeCalculated.overTime.content
                        = (int)SiGlobalVars.Instance.DEFAULT_WORKING_HOURS;

                    workTimeCalculated.workTimeHours.content = 0;
                }

            }
            else
            {
                //ie. not friday

                if (mepOvertime.overtime == SiGlobalVars.Instance.ABSENT ||
               thisDate_is_fridayOrHoliday == false
               )
                {
                    //ie not firday and absent

                    workTimeCalculated.workTimeHours.content = 0;
                    workTimeCalculated.overTime.content = 0;
                }
                else
                {
                    //ie not absent and not friday
                    //means there should be some overtime
                    int resultOvertime;
                    if (int.TryParse(mepOvertime.overtime, out resultOvertime) == false)
                    {
                        MessageBox.Show("The overtime value for Employee = " +
                            mepStyleWrap.code + " for the date = "
                            + mepOvertime.date.ToString() + "; Cell Address = "
                            + mepOvertime.fullCell.Address);
                        return null;
                    }
                    else
                    {
                        workTimeCalculated.workTimeHours.content
                            = SiGlobalVars.Instance.DEFAULT_WORKING_HOURS;
                        workTimeCalculated.overTime.content
                            = SiGlobalVars.Instance.DEFAULT_WORKING_HOURS
                            + resultOvertime;
                    }


                }

            }


            return workTimeCalculated;
        }

        private static Boolean write_data_to_payLoadFormat_from_mepStyle
       (PayLoadWrap.Day.Employee payLoadWrapDayEmpl,
       MepStyleWrap mepStyleWrap,
       Boolean thisDate_is_fridayOrHoliday)
        {
            payLoadWrapDayEmpl.job_siteNo.content
               = mepStyleWrap.siteNo.content;

            foreach (var dateOverTime in mepStyleWrap.dateOvertimes)
            {
                WorkTimeCalculatedWarp workTimeCalculated =
                MixTimeSheetHandler.calculate_workTime_from_MEPStyleOvertime
                (dateOverTime,
                thisDate_is_fridayOrHoliday, mepStyleWrap);

                if (workTimeCalculated == null)
                    return false;

                //now we need to write it to payload

                //for worktime
                payLoadWrapDayEmpl.workTime.content 
                    = workTimeCalculated.workTimeHours.content;
                payLoadWrapDayEmpl.workTime.contentInStr 
                    = workTimeCalculated.workTimeHours.content.ToString();
                payLoadWrapDayEmpl.workTime.fullCell.Value 
                    = payLoadWrapDayEmpl.workTime.contentInStr;
                //for overtime
                payLoadWrapDayEmpl.overTime.content 
                    = workTimeCalculated.overTime.content;
                payLoadWrapDayEmpl.overTime.contentInStr 
                    = workTimeCalculated.overTime.content.ToString();
                payLoadWrapDayEmpl.overTime.fullCell.Value
                    = payLoadWrapDayEmpl.overTime.contentInStr;
            }


            return true;

        }

        private static Boolean write_data_to_payLoadFormat_from_MultiTrans
            (PayLoadWrap.Day.Employee payLoadWrapDayEmpl,
            MultiTransWrap multiWrap,
            Boolean thisDate_is_fridayOrHoliday,
            bool printBio_inPayLoad)
        {

            payLoadWrapDayEmpl.job_siteNo.content
                = multiWrap.siteNoMechFormat.shortName.content;


            WorkTimeCalculatedWarp workTimeCalculated =
                PayLoadHelper.Calculate_worktime_from_bioTotalWorkTime(multiWrap.totalTimeWorked, thisDate_is_fridayOrHoliday);

            if (workTimeCalculated == null)
                return false;

            //now we need to write it to payload

            //for worktime
            payLoadWrapDayEmpl.workTime.content = workTimeCalculated.workTimeHours.content;
            payLoadWrapDayEmpl.workTime.contentInStr = workTimeCalculated.workTimeHours.content.ToString();
            payLoadWrapDayEmpl.workTime.fullCell.Value = payLoadWrapDayEmpl.workTime.contentInStr;
            //for overtime
            payLoadWrapDayEmpl.overTime.content = workTimeCalculated.overTime.content;
            payLoadWrapDayEmpl.overTime.contentInStr = workTimeCalculated.overTime.content.ToString();
            payLoadWrapDayEmpl.overTime.fullCell.Value = payLoadWrapDayEmpl.overTime.contentInStr;

            //now sometimes for testing purpose 
            //we can print the biometric acutal time from the multiple transactions 
            //to payload excel . this is to cross check if the caluclated overtime is 
            //correct or not
            if (printBio_inPayLoad == true)
            {
                var bioTimeCell = payLoadWrapDayEmpl.overTime.fullCell.Next;
                bioTimeCell.Value = multiWrap.totalTimeWorked.contentInString;

            }

            return true;

        }

        internal Boolean correct_siteNo_in_multiTrans_with_shortMechSiteNo
            (List<MultiTransWrap> multiTransWraps)
        {
            Boolean not_even_one_SiteNo_is_corrected = true;
            foreach (var item in multiTransWraps)
            {
                //first read or get shortnames
                //first get the shortname
                if (item.siteNo == null)
                    continue;   //sometimes siteno is not available then ignore the cell for correction

                String shortname =
                    CommonOperations.convert_siteNo_to_SiteNoMechFormat_ShortName(item.siteNo);

                if (shortname == null)
                {
                    MessageBox.Show("Failed to correct Site No. Process is Aborted; No changes have " +
                        "been made in the multiTrans Excel File");
                    return false;

                }
                item.siteNoMechFormat.shortName.content
                    = shortname;

            }
            foreach (var item in multiTransWraps)
            {
                if (item.siteNo == null)
                    continue;   //sometimes siteno is not available then ignore the cell for correction
                //now write 
                item.siteNo.fullCell.Value = item.siteNoMechFormat.shortName.content;
                item.siteNoMechFormat.shortName.fullCell = item.siteNo.fullCell;
                not_even_one_SiteNo_is_corrected = false;

            }

            if (not_even_one_SiteNo_is_corrected == true)
                return false;

            return true;
        }
    }
}
