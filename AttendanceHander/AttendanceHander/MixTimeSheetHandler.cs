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
            List<String> skipList_Emp_no = new List<String>();

            foreach (var multiWrap in multiTransWraps)
            {
                if (skipList_Emp_no.Contains(multiWrap.personnelNo.content))
                    continue; //if skip list contains this employee no, then skip

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
                    if (multiWrap.siteNoMechFormat == null)
                        multiWrap.siteNoMechFormat = new MultiTransWrap.SiteNoMechFormat();
                    if (multiWrap.siteNoMechFormat.longSiteNo == null)
                        multiWrap.siteNoMechFormat.longSiteNo = new StrItemWrap();
                    if (multiWrap.siteNoMechFormat.shortName == null)
                        multiWrap.siteNoMechFormat.shortName = new StrItemWrap();

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
                        .compare_multiTrans_employeeNo_to_MepStyle_employeeNo
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

            }


            return true;
        }

        internal Boolean correct_siteNo_in_multiTrans_with_shortMechSiteNo
            (List<MultiTransWrap> multiTransWraps)
        {
            foreach(var item in multiTransWraps)
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

            }
            

            return true;
        }
    }
}
