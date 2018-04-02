using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class WordTransform
    {
        public static string RemoveSpecialCharacters(string word)
        {
            return Regex.Replace(word, @"[:!@#$%^&*()}{|\"":?><\[\]\\;'\.,~0-9]", "");
        }
    }
}
