using System.Text.RegularExpressions;

namespace CommentsAnalysis
{
    class WordTransform
    {
        public static string RemoveSpecialCharacters(string word)
        {
            return Regex.Replace(word, @"[:!@#$%^&*()}{|\"":?><\[\]\\;'\.,~0-9]", "");
        }
    }
}
