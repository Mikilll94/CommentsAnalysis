using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class General
    {
        public static string RemoveSpecialCharacters(string word)
        {
            return Regex.Replace(word, @"[:!@#$%^&*()}{|\"":?><\[\]\\;'\.,~0-9]", "");
        }
    }
}
