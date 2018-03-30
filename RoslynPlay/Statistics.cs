using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public class Statistics
    {
        public string CommentLocation { get; }
        public string MethodName { get; }
        public int? WordsCount { get; }
        public bool? HasNothing { get; }
        public bool? HasQuestionMark { get; }
        public bool? HasExclamationMark { get; }
        public bool? HasCode { get; }
        public double? CoherenceCoefficient { get; }

        public Statistics(string content, int lineEnd, string type, CommentLocationStore commentLocationstore)
        {
            char[] delimiters = new char[] { ' ', '\t', '\r', '\n' };
            if (type == "single_line_comment")
            {
                WordsCount = content.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            if (commentLocationstore.CommentLocations.ContainsKey(lineEnd))
            {
                CommentLocation = commentLocationstore.CommentLocations[lineEnd][0];
                MethodName = commentLocationstore.CommentLocations[lineEnd][1];
            }

            HasNothing = new Regex("nothing", RegexOptions.IgnoreCase).IsMatch(content);
            HasQuestionMark = new Regex(@"\?").IsMatch(content);
            HasExclamationMark = new Regex("!").IsMatch(content);

            int linesWithCode = 0;
            int linesCount = 0;

            string[] instructionsKeywords = new string[] { "if", "for", "try", "catch", "while", "do" };
            string instructionsRegexPart = "";
            foreach (var keyword in instructionsKeywords)
            {
                instructionsRegexPart += (keyword + @"\s?\(.*\)|");
            }
            Regex hasCodeRegex = new Regex(@"(" + instructionsRegexPart + "[;{}]|void|==|!=)");

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
            HasCode = (double)linesWithCode / linesCount > 0.1;

            if (MethodName != null && CommentLocation == "method_description")
            {
                string[] methodNameWords = Regex.Replace(MethodName, "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))", "$1 ").Split(" ");
                methodNameWords = methodNameWords.Where(word => word.Length > 2).ToArray();

                content = Regex.Replace(content, "[.,]", "");
                string[] contentWords = content.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                contentWords = contentWords.Where(word => word.Length > 2).ToArray();
                CoherenceCoefficient = RoslynPlay.CoherenceCoefficient.Compute(contentWords, methodNameWords);
            }
        }

        public bool? WordsCountBad()
        {
            if (WordsCount == null || WordsCount == 0)
            {
                return null;
            }
            else
            {
                return WordsCount <= 2;
            }
        }

        public bool? CoherenceCoefficientBad()
        {
            if (CoherenceCoefficient == null)
            {
                return null;
            }
            else
            {
                return CoherenceCoefficient == 0 || CoherenceCoefficient > 0.5;
            }
        }

    }
}
