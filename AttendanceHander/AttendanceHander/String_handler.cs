using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace AttendanceHander
{
    class String_handler
    {
       
        public All_const.str_type is_this_string_alpha_numeric_or_numeric_or_alpha_only(String string_to_check)
        {
            Boolean contains_digits = false;
            Boolean contains_alphabets = false;
            if (string_to_check == null)
                return All_const.str_type.Null;
            char[] stream_of_char = string_to_check.ToCharArray();
            foreach (char this_char in stream_of_char)
            {
                if (char.IsDigit(this_char))
                    contains_digits = true;
                else if (char.IsLetter(this_char))
                    contains_alphabets = true;
            }

            if (contains_digits == true && contains_alphabets == true)
                return All_const.str_type.Alphanumeric;
            else if (contains_digits == true && contains_alphabets == false)
                return All_const.str_type.Numeric;
            else  
                return All_const.str_type.Alpha_only;
        }


        public Double similar_word_percent(String current_word,
          String word_to_check, Boolean case_sensitive)
        {

            if (current_word == null && word_to_check == null)
                return 100;
            else if (current_word == null && word_to_check != null)
                return 0;
            else if (current_word != null && word_to_check == null)
                return 0;
            else if (current_word == word_to_check)
                return 100;
            else if (case_sensitive == false &&
                string.Equals(current_word, word_to_check, StringComparison.CurrentCultureIgnoreCase) == true)
                return 100;

            //if none of the above are true means
            //words are not exactly equal
            //but check the chances it is very similiar; that is more than 80% etc
            int total_char_count;
            int total_matching_chars = 0;
            int total_range_of_loop = 0;
            Char[] curr_word_chars = current_word.ToCharArray();
            Char[] word_to_check_chars = word_to_check.ToCharArray();
            if (curr_word_chars.Length > word_to_check_chars.Length)
                total_char_count = curr_word_chars.Length;
            else
                total_char_count = word_to_check_chars.Length;

            if (curr_word_chars.Length < word_to_check_chars.Length)
                total_range_of_loop = curr_word_chars.Length;
            else
                total_range_of_loop = word_to_check_chars.Length;

            if (case_sensitive == false)
            {
               
                for (int i = 0; i < total_range_of_loop; i++)
                {
                    if (Char.ToLower(curr_word_chars[i]) ==
                        Char.ToLower(word_to_check_chars[i]))
                    {
                        total_matching_chars = total_matching_chars + 1;
                    }

                }
            }
            else
            {
                for (int i = 0; i < total_range_of_loop; i++)
                {
                    if (curr_word_chars[i] ==
                        word_to_check_chars[i])
                    {
                        total_matching_chars = total_matching_chars + 1;
                    }

                }
            }
            Double result=0;


            result = (Double) total_matching_chars/total_char_count * 100;
            return result;

        }
        public Boolean similar_words(String current_word,
            String word_to_check,Boolean case_sensitive, float threashold_percent)
        {

            if (current_word == null && word_to_check == null)
                return true;
            else if (current_word == null && word_to_check != null)
                return false;
            else if (current_word != null && word_to_check == null)
                return false;
            else if (current_word == word_to_check)
                return true;
            else if (case_sensitive == false &&
                string.Equals(current_word, word_to_check, StringComparison.CurrentCultureIgnoreCase) == true)
                return true;

            //if none of the above are true means
            //words are not exactly equal
            //but check the chances it is very similiar; that is more than 80% etc
            float total_char_count;
            float total_matching_chars = 0;
            Char[] curr_word_chars = current_word.ToCharArray();
            Char[] word_to_check_chars = word_to_check.ToCharArray();
            if (curr_word_chars.Length > word_to_check_chars.Length)
                total_char_count = curr_word_chars.Length;
            else
                total_char_count = word_to_check_chars.Length;


            int total_range_of_loop = 0;
            if (curr_word_chars.Length < word_to_check_chars.Length)
                total_range_of_loop = curr_word_chars.Length;
            else
                total_range_of_loop = word_to_check_chars.Length;

            if (case_sensitive == false)
            {
                for(int i=0; i< total_range_of_loop; i++)
                {
                    if(Char.ToLower(curr_word_chars[i])==
                        Char.ToLower(word_to_check_chars[i]))
                    {
                        total_matching_chars = total_matching_chars + 1;
                    }

                }
            }
            else
            {
                for (int i = 0; i < total_range_of_loop; i++)
                {
                    if (curr_word_chars[i] ==
                        word_to_check_chars[i])
                    {
                        total_matching_chars = total_matching_chars + 1;
                    }

                }
            }

            float result;
            result = (float) total_matching_chars / total_char_count * 100;
            if (result >= threashold_percent)
                return true;
            else
                return false;

        }
        public Nullable<DateTime> convert_string_to_date_time(String in_date)
        {
            if (in_date == null)
                return null;
            DateTime oDate = new DateTime();
            Boolean conversion_success = DateTime.TryParse(in_date,out oDate);
            if (conversion_success == false)
                return null;

            //so return 
            return oDate;
          

        }
        public String[] get_all_words_in_a_sentence(String current_sentence,
            Char[] word_separating_char)
        {
            String[] all_words;
            
           all_words =  current_sentence.Split(word_separating_char,StringSplitOptions.RemoveEmptyEntries);
            return all_words;
        }

        public Boolean similar_sentences(String current_sentence,
            String sentence_to_check_with, Boolean check_word_order,
            Boolean case_sensitive,
            float word_threshold_percent,
            float sentence_threshold_percent,
            Boolean sentence_words_count_should_be_same)
        {
            float total_words_count = 0;
            Dictionary<String, float> results  = new Dictionary<string, float>();

            if (current_sentence == null && sentence_to_check_with == null)
                return true;
            else if (current_sentence == null && sentence_to_check_with != null)
                return false;
            else if (current_sentence != null && sentence_to_check_with == null)
                return false;

                String[] curr_sentence_word=null;
            char[] word_separating_chars = { ' ', '-', '.' };

            curr_sentence_word = get_all_words_in_a_sentence(
                current_sentence, word_separating_chars);
         
            total_words_count = curr_sentence_word.Length;
            String[] sentence_to_check_word = null;
            if (sentence_to_check_with != null)
                curr_sentence_word = sentence_to_check_with.Split(' ');
            
            foreach(String word in curr_sentence_word)
            {
                results.Add(word, 0);
            }

            if (check_word_order == true)
            {
                for(int i =0; i< curr_sentence_word.Length;
                    i++)
                {
                    if(similar_words(curr_sentence_word[i],
                        sentence_to_check_word[i],case_sensitive,
                        word_threshold_percent))
                    {
                        results[curr_sentence_word[i]]= 100;
                    }
                }
            }
            else
            {
                for (int i = 0; i < curr_sentence_word.Length;
                    i++)
                {
                    foreach(String word in sentence_to_check_word)
                    {
                        if (similar_words(curr_sentence_word[i],
                        word, case_sensitive,
                        word_threshold_percent))
                        {
                            results[curr_sentence_word[i]] = 100;

                        }
                      
                    }
                }
            }
            float total_result_sum=0;
            float total_out_of=0;
            if (sentence_words_count_should_be_same==false)
                total_out_of = curr_sentence_word.Length * 100;
            else
            {
                if (curr_sentence_word.Length >= sentence_to_check_word.Length)
                    total_out_of = curr_sentence_word.Length * 100;
                else
                    total_out_of = sentence_to_check_word.Length*100;
            }
           
            //now check the result
            foreach(KeyValuePair<String,float> result in results)
            {
                total_result_sum = total_result_sum + result.Value;
            }

            float result_value = total_result_sum / total_out_of * 100;

            if (result_value >= sentence_threshold_percent)
                return true;
            else
                return false;
        }
    }
}
