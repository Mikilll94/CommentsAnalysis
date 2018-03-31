using System.IO;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public static class CodeDetector
    {
        public static bool HasCode(string content)
        {
            int linesWithCode = 0;
            int linesCount = 0;

            string[] instructionsKeywords = new string[] { "if", "for", "try", "catch", "while", "do" };
            string instructionsRegexPart = "";
            foreach (var keyword in instructionsKeywords)
            {
                instructionsRegexPart += (keyword + @"\s?\(.*\)|");
            }
            Regex hasCodeRegex = new Regex(instructionsRegexPart + @"^{|^}|\bvoid\b|\bvar\b|==|!=|;$|^\/\/");

            using (var reader = new StringReader(content))
            {
                for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
                {
                    if (hasCodeRegex.IsMatch(line))
                    {
                        linesWithCode++;
                    }
                    linesCount++;
                }
            }
            return (double)linesWithCode / linesCount > 0.1;
        }
    }
}
