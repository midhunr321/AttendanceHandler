using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Diagnostics;





namespace AttendanceHander
{
    public class EXCEL_HELPER
    {
        //LAST UPDATED ON 18-OCTOBER 2019

        public Excel.Worksheet worksheet;
        public EXCEL_HELPER(Excel.Worksheet worksheet_arg)
        {

            worksheet = worksheet_arg;
            //CONSTRUCTOR

        }

        public static void hide_unhide_excel_row(ref Excel.Range fullCell, Boolean hidden)
        {
            fullCell.EntireRow.Hidden = hidden;
        }
        public Excel.Range get_lowest_column_cell_from_search_result(List<Excel.Range> searchResults)
        {
            if (searchResults == null)
                return null;
            Excel.Range lowest_column_result = null;
            foreach (Excel.Range result in searchResults)
            {
                if (lowest_column_result == null)
                    lowest_column_result = result;

                if (result.Column < lowest_column_result.Column)
                    lowest_column_result = result;
            }
            return lowest_column_result;
        }
        public Excel.Range get_largest_column_cell_from_search_result(List<Excel.Range> searchResults)
        {
            if (searchResults == null)
                return null;

            Excel.Range largest_column_result = null;
            foreach (Excel.Range result in searchResults)
            {
                if (largest_column_result == null)
                    largest_column_result = result;

                if (result.Column > largest_column_result.Column)
                    largest_column_result = result;
            }
            return largest_column_result;
        }
        public void turnOff_filters(Excel.Worksheet worksheet)
        {
            if (worksheet.AutoFilter != null && worksheet.AutoFilterMode == true)
            {
                worksheet.AutoFilter.ShowAllData();
            }
        }

        public Excel.Range return_immediate_below_cell
            (Excel.Range nonMergedSingleCell)
        {
            int currentrow = nonMergedSingleCell.Row;
                    
            int nextRowNo = currentrow + 1;
            Excel.Range beneathCell = worksheet.Cells[nextRowNo, 
                nonMergedSingleCell.Column];
            return beneathCell;
        }
        public Excel.Range return_next_adjacent_range(Excel.Range current_range)
        {
            int current_range_merged_columns = current_range.Columns.Count;
            int current_range_row_no = current_range.Row;
            int current_range_col_no = current_range.Column;

            //so next_cell column
            int next_cell_required_col = current_range_col_no
                + current_range_merged_columns;

            //now carryout loop



            Excel.Range next_cell = current_range.Next;
            while (next_cell.Column != next_cell_required_col)
            {
                if (next_cell.Column > next_cell_required_col)
                    return null;

                next_cell = return_next_adjacent_range(
                    next_cell);
            }

            return (next_cell);

        }
        public int get_last_column_no_of_a_merge_cell(Excel.Range fullCell)
        {
            if (fullCell == null)
                return -1;

            if (fullCell.MergeCells ==true)
            {
                //ie contains Merge Cells
                int last_column_no=0;
                foreach(Excel.Range cell in fullCell.Cells)
                {
                    if (cell.Column > last_column_no)
                        last_column_no = cell.Column;
                }
                return last_column_no;
            }
            else
            {
                return (fullCell.Column);
            }
        }

        public Boolean is_this_a_merged_cell(Excel.Range fullCell)
        {
            if (fullCell.MergeCells==true)
                return true;
            
            return false;
        }
        private Dictionary<Excel.Range, String> get_cell_with_address(List<Excel.Range> list_of_cell)
        {
            string complete_address;
            Dictionary<Excel.Range, String> cells_with_address
                = new Dictionary<Excel.Range, string>();
            foreach (Excel.Range cell in list_of_cell)
            {
                complete_address = cell.MergeArea.Address;
                cells_with_address.Add(cell, complete_address);

            }
            return cells_with_address;

        }
        public List<String> return_same_address_cells(
            List<List<Excel.Range>> list, int no_of_repeatings_required)
        {

            var list_of_address = from each_list in list
                                  from range in each_list
                                  select range.MergeArea.Address;

            var grouped_addresses = from address in list_of_address
                                    group address by address into g
                                    select new { g };

            //to select the largest count;
            int largest_count = 0;

            List<String> list_of_most_repeated_cell_addresses =
                new List<string>();
            foreach (var group_ in grouped_addresses)
            {
                if (group_.g.Count() > largest_count)
                {
                    largest_count = group_.g.Count();
                }

            }
            //second iteration is to cross check if there is any tie
            //for example if the largest_count = 2 but
            //there are more than one group with count 2
            //then which group should be selected?
            //we can't simply ignore the other result
            //so we need to consider the tie as well
            //so for this application lets carryout second iteration
            foreach (var group_ in grouped_addresses)
            {
                if (group_.g.Count() == largest_count)
                {

                    list_of_most_repeated_cell_addresses.Add(group_.g.Key);
                }

            }



            return list_of_most_repeated_cell_addresses;
        }
        public class Match_percent_cell
        {
            public Excel.Range cell_;
            public Double match_percent;

        }

        private List<Excel.Range> compare_the_search_result_with_source(String source_search_string
            , List<String> cell_addresses_of_search_result,
            Boolean ignore_special_chars)
        {
            Dictionary<String, Excel.Range> cell_dic = new Dictionary<string, Excel.Range>();
            foreach (String cell_address in cell_addresses_of_search_result)
            {
                Excel.Range cell_ = worksheet.Range[cell_address];
                String cell_value = get_value_of_merge_cell(cell_);
                if (ignore_special_chars == true)
                {
                    String[] all_words =
                        get_all_words_of_senten_without_punct(cell_value);
                    String merged_string = join_words_to_sentence(all_words);
                    cell_value = merged_string;
                }
                cell_dic.Add(cell_value, cell_);

            }

            //now that we initiated the cell_dic 
            //lets compare it with source_search_string

            StringHandler string_Handler =
                    new StringHandler();

            List<Match_percent_cell> match_percent_list =
                new List<Match_percent_cell>();

            foreach (KeyValuePair<String, Excel.Range> item in cell_dic)
            {
                Double matching_percent = string_Handler
                    .similar_word_percent(item.Key, source_search_string
                    , false);
                Match_percent_cell match_Percent_Cell = new Match_percent_cell();
                match_Percent_Cell.cell_ = item.Value;
                match_Percent_Cell.match_percent = matching_percent;
                match_percent_list.Add(match_Percent_Cell);

            }

            //now return the larger percent
            Double largest_percent = 0;
            List<Excel.Range> search_results = new List<Excel.Range>();
            foreach (Match_percent_cell item in match_percent_list)
            {
                if (item.match_percent > largest_percent)
                {
                    largest_percent = item.match_percent;

                }
            }

            //we need one more iteration same as above
            //this is because if there are more than one search result with
            //same largest_percent
            //it needs to be considered.
            foreach (Match_percent_cell item in match_percent_list)
            {
                if (largest_percent == 0)
                    break;
                if (item.match_percent == largest_percent)
                {
                    search_results.Add(item.cell_);

                }
            }




            if (largest_percent >= 80)
                return search_results;
            else
                return null;

        }
        private String[] get_all_words_of_senten_without_punct(String source_string)
        {
            if (source_string == null)
                return null;

            StringHandler string_Handler = new StringHandler();

            char[] word_separating_chars = { ' ', '-', '.' };
            return (string_Handler.get_all_words_in_a_sentence(
                source_string, word_separating_chars));

        }
        private String join_words_to_sentence(String[] all_words)
        {
            if (all_words == null)
                return null;
            int i = 0;
            String merged_string = null;
            foreach (String word in all_words)
            {
                if (i == 0)
                    merged_string = merged_string +
                        word;
                else
                    merged_string = merged_string
                        + " " + word;

                i++;
            }

            return merged_string;
        }
        private List<Excel.Range> search_smartly_by_similarity_check(
            String search_string,
            Excel.XlSearchOrder xlSearchOrder, Excel.XlSearchDirection xlSearchDirection,
            Boolean matchcase = false)
        {

            String[] all_search_words;

            all_search_words = get_all_words_of_senten_without_punct(search_string);

            // similary checking procedure
            // search for each words in the excel
            // get all cells for each word
            //store these cells in dictionary
            //now check cell address of each cell
            // the most common cell address means
            // the words are in the   same cell

            Dictionary<String, List<List<Excel.Range>>> search_results_dic =
                new Dictionary<string, List<List<Excel.Range>>>();

            List<List<Excel.Range>> s_result_for_word = new List<List<Excel.Range>>();
            foreach (String word in all_search_words)
            {
                var s = search_for_cell(word, xlSearchOrder, xlSearchDirection,
                    matchCase:matchcase);
                if (s != null)
                    s_result_for_word.Add(s);
                search_results_dic.Add(word, s_result_for_word);
            }
            int word_count = all_search_words.Length;
            int minimum_required_repetition = (int)Math.Round(0.8 * word_count, 0);
            List<String> search_results_with_common_cell_address;
            search_results_with_common_cell_address =
                return_same_address_cells(s_result_for_word,
                minimum_required_repetition);

            //now there is a problem
            //eg: for search word "S No."
            //we first search for 'S' ; which returns say like 87 results
            //then search for "No." which indeed may result say like 3
            //then we get to cells with content as
            //1. S No
            //2. Site No.
            //so inorder to filter errors like this lets
            //check the order of the letters as well
            List<Excel.Range> search_results;

            //now get all words without punctuations together in a
            //string for carrying out comparision
            String search_str_without_punctua = null;
            search_str_without_punctua = join_words_to_sentence(all_search_words);

            search_results =
              compare_the_search_result_with_source(search_str_without_punctua,
              search_results_with_common_cell_address, true);


            if (search_results != null)
                return search_results;
            else
                return null;



        }


        public List<Excel.Range> find_fix_column_heading(String table_col_name,
            Excel.XlSearchDirection xlSearchDirection,
            Excel.XlSearchOrder xlSearchOrder,
            Boolean matchcase)
        {
            List<Excel.Range> sresult = new List<Excel.Range>();
            sresult = search_for_cell(table_col_name,
                xlSearchOrder,xlSearchDirection,matchcase);
            if (sresult.Count == 0)
                sresult = search_smartly_by_similarity_check(table_col_name,
                    xlSearchOrder, xlSearchDirection, matchcase);
            //if have more than 1 search result for table column name
            // then what we will do?
            // an idea...search for rest of the headings.
            //if majority of the heading is in one particular raw number
            // that means that row should be heading.
            //so if more than one search results then
            // we can filter it out.
            List<Excel.Range> heading_cell = new List<Excel.Range>();
            //to find the top most cell; we find the lowest row no & cell no address cell
            List_helper_for_excel list_Helper_For_Excel =
                new List_helper_for_excel();
            //so even if more than one search result lets keep it for time being
            heading_cell = sresult;
            return heading_cell;
        }

        public Excel.Range return_top_search_range(List<Excel.Range> search_list)
        {
            if (search_list == null || search_list.Count == 0)
                return null;
            Excel.Range top_most_search_result = search_list[0];//assume first
            int lowest_row = search_list[0].Row;//assume first result is top most
            foreach (Excel.Range item in search_list)
            {
                if (item.Row < lowest_row)
                    top_most_search_result = item;
            }

            return top_most_search_result;
        }

        public int insert_text_infront_in_cells(Excel.Range cell_range, String text_to_insert)
        {
            int count = 0;
            foreach (Excel.Range cell in cell_range)
            {
                String filled_content = cell.Value2;
                cell.Value2 = text_to_insert + " " + filled_content;
                count++;

            }

            return count;
        }
        public string get_value_of_merge_cell(Excel.Range merge_cell)
        {
            String cell_value = "";
            foreach (Excel.Range cell_ in merge_cell)
                cell_value = cell_value + cell_.Value2;

            return cell_value;

        }
        public Excel.Range return_full_merg_cell(Excel.Range part_of_merge_cell)
        {

            if (part_of_merge_cell.MergeCells)
            {
                string complete_address;
                complete_address = part_of_merge_cell.MergeArea.Address;
                Excel.Range complete_cell_range;
                complete_cell_range = worksheet.Range[complete_address];
                return complete_cell_range;
            }
            else
            {
                return part_of_merge_cell;
            }
        }
        public void change_cell_interior_color(ref Excel.Range cell_, System.Drawing.Color color)
        {
            if (cell_ != null)
                cell_.Interior.Color = color;

        }

        public Boolean cells_are_in_the_same_row (List<Excel.Range> Cells)
        {
            int expected_row = Cells[0].Row;

            foreach(Excel.Range cell in Cells)
            {
                if (cell.Row != expected_row)
                    return false;
            }

            return true;
        }
        private List<Excel.Range> search_for_cell(String find_text,
            Excel.XlSearchOrder xlSearchOrder,
            Excel.XlSearchDirection xlSearchDirection,
            Boolean matchCase )
        {
            //Get the used Range
            Excel.Range usedRange = worksheet.UsedRange;
            //Iterate though the rows and find the text is there in the row or not.
            //if the text is there that means it is the row.
            //thus we get the coloumn no. and row no.
            //if it is a table, then this row should be heading
            // all data coming under this row(heading) is our data of interest.


            //Ex. Iterate through the row's data and put in a string array

            List<Excel.Range> sresult = new List<Excel.Range>();
            //for first result.
            //first search result is stored as sresult[0]
            //start search from first cell

            Excel.Range current_find = usedRange.Find(What: find_text,
               SearchOrder: xlSearchOrder,SearchDirection: xlSearchDirection,
               MatchCase:matchCase);

            if (current_find != null)
                sresult.Add(current_find);



            int i = 0;
            int total_search_count = sresult.Count;

            do
            {
                if (total_search_count == 0)
                    break;
                if (sresult[i] != null)
                {
                    i++;
                    //i.e we already got one result. now check if there is any next one.

                    Excel.Range temp_result = usedRange.FindNext(sresult[i - 1]);
                    //bcz to findnext from the last object ..last object is sresult[i-1]
                    //now adding next find to list make sure the following:
                    //1. found one is not empty
                    //2. the found result again repeated . ie if it is the same
                    //first value or not.
                    if (temp_result != null && temp_result.Address != sresult[0].Address)
                        sresult.Add(temp_result);
                    else
                        return sresult;

                }
            } while (sresult[i] != null && sresult[i].Address != sresult[0].Address);


            return sresult;

        }



    }

}
