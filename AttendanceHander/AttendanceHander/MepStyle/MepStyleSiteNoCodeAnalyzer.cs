﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
    class MepStyleSiteNoCodeAnalyzer
    {

        class Codewrap
        {
            public int StartIndex;
            public String text;
            public int EndIndex;

            public Codewrap(int startIndex, string text, int endIndex)
            {
                StartIndex = startIndex;
                this.text = text;
                EndIndex = endIndex;
            }
        }

        private Boolean find_start_of_transfer_date(String transferCode,
            Codewrap to, out String transferStartDate)
        {
            transferStartDate = null;
            //now find the start date of plumber transfer
            //start date of site transfer is before "to"

            //now check how may digits for transfer start date is available
            int length_of_transfer_start_date = to.StartIndex + 1;
            //+1 because index always start from 0; but length is the magnitude

            if (length_of_transfer_start_date > 2)
                return false; //becase date from 1 to 31 is only 2 digits


            String startDAte = transferCode.Substring(0,
                length_of_transfer_start_date);


            int parsedTransferStartDate;

            if (invalidate_transfer_date(startDAte,
                 out parsedTransferStartDate)
                 == false)
                return false;

            transferStartDate = startDAte;
            return true;
        }

        private Codewrap find_TO(String transferCode)
        {
            //first check if the string is in the code format
            //an example format is 10to10_M265;

            //first find "to"
            StringHandler stringHandler = new StringHandler();

            int toStartIndex;
            int toEndIndex;
            if (stringHandler.start_end_index_of_substring(
                  transferCode, "to", out toStartIndex, out toEndIndex)
                  == false)
                return null;

            Codewrap to = new Codewrap(toStartIndex, "to", toEndIndex);
            return to;

        }

        private Codewrap find_underscore(String transferCode,
            Codewrap to)
        {
            StringHandler stringHandler = new StringHandler();
            //now find the underscore "_"
            int startindex_;
            int endIndex_;
            if (stringHandler.start_end_index_of_substring(
                  transferCode, "_", out startindex_, out endIndex_)
                  == false)
                return null;

            Codewrap underscore = new Codewrap(to.StartIndex, "_", to.EndIndex);
            return underscore;
        }
        public Boolean analyze_string(String transferCode)
        {
            //first check if the string is in the code format
            //an example format is 10to10_M265;

            //first find "to"
            StringHandler stringHandler = new StringHandler();


            Codewrap to = find_TO(transferCode);
            if (to == null)
                return false;

            //now find the underscore "_"

            Codewrap underscore = find_underscore(transferCode, to);
            if (underscore == null)
                return false;

            //now find the start date of plumber transfer
            //start date of site transfer is before "to"
            String transferStartDate;

            if (find_start_of_transfer_date(transferCode, to, out transferStartDate)
                 == false)
                return false;


            //now find the end date of plumber site shift 
            String transferEndDate;
            if (find_end_date_of_site_transfer(transferCode,
                 to, underscore, out transferEndDate)
                 == false)
                return false;

            //now to get the site no "eg. M273"
            //Site no is after the underscore

            int sitenoStartIndex = underscore.EndIndex + 1;
            String siteno = transferCode.Substring(sitenoStartIndex);
            if (invalidate_siteNo(siteno)
                == false)
                return false;

        }

        private Boolean invalidate_siteNo(String siteno)
        {
            StringHandler stringHandler = new StringHandler();
            if (stringHandler
                 .is_this_string_alpha_numeric_or_numeric_or_alpha_only(siteno)
                 != All_const.str_type.Alphanumeric)
                return false;

            if (siteno.Length > 5)
                return false;

            return true;

        }
        private Boolean find_end_date_of_site_transfer(String transferCode,
            Codewrap to, Codewrap underscore, out String transferEndDate)
        {
            transferEndDate = null;
            //now find the end date of plumber site shift 

            int start_index = to.EndIndex;
            int end_index = underscore.StartIndex;
            int length = (start_index - end_index) - 1;

            int start_index_after_TO = to.EndIndex + 1;
            String EndDAte = transferCode.Substring(start_index_after_TO,
               length);

            int parsedTransferEndDate;
            if (invalidate_transfer_date(EndDAte,
                 out parsedTransferEndDate)
                 == false)
                return false;

            transferEndDate = EndDAte;
            return true;
        }
        public Boolean invalidate_transfer_date(String transferDate,
            out int parsedTransferDate)
        {
            parsedTransferDate = -1;
            StringHandler stringHandler = new StringHandler();

            if (stringHandler
                 .is_this_string_alpha_numeric_or_numeric_or_alpha_only
                 (transferDate)
                 != All_const.str_type.Numeric)
                return false;

            int convertedDate;
            //now check if the start date is actually a date or not
            if (int.TryParse(transferDate, out convertedDate)
                 == false)
                return false; //ie trafer date is not a number

            //now check if the date is more than 31
            if (convertedDate > 31)
                return false; //means its not a valid date

            parsedTransferDate = convertedDate;
            return true;
        }


    }
}
