using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceHander
{
    class SearchHelper
    {

        public Excel.Range searchTextWithoutPunctSpaceCase(String sourceString)
        {


        }

        
        private String extract_string_without_punct_and_space(String sourceString)
        {
            StringBuilder stringBuilder = new StringBuilder();
            String strWithoutSpecialChar;
            foreach (Char character in sourceString)
            {
                if (!Char.IsPunctuation(character) &&
                    !Char.IsWhiteSpace(character) &&
                   !Char.IsSymbol(character) &&
                    !Char.IsSeparator(character))
                {
                    stringBuilder.Append(character);
                }
            }

            strWithoutSpecialChar = stringBuilder.ToString();
            return strWithoutSpecialChar;
        }


    }
}
