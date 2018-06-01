using System.IO;
using System.Text.RegularExpressions;

namespace CommentsAnalysis
{
    public static class CodeDetector
    {
        public static bool HasCode(string content)
        {
            int linesWithCode = 0;
            int linesCount = 0;

            string codeKeywords = "";

            string[] instructionKeywords = new string[] { "if", "for", "try", "catch", "while", "do", "typeof", "nameof" };
            foreach (var keyword in instructionKeywords)
            {
                codeKeywords += $@"{keyword}\s?\(.*\)|";
            }

            string[] expressionKeywords = new string[] { "void", "var" };
            foreach (var keyword in expressionKeywords)
            {
                codeKeywords += $@"\b{keyword}\b|";
            }

            Regex hasCodeRegex = new Regex(codeKeywords + @"^{|^}|==|!=|;$|^\/\/");

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
