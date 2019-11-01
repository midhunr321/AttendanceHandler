using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using AttendanceHander.DailyTransactions;
using AttendanceHander.MultipleTransaction;
using System.Windows.Forms;

namespace AttendanceHander
{
    class MixTimeSheetHandler
    {

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
