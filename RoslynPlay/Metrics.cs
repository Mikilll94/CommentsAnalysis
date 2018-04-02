using System;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public class Metrics
    {
        public string CommentLocation { get; }
        public string MethodName { get; }
        public int? WordsCount { get; }
        public bool? HasNothing { get; }
        public bool? HasQuestionMark { get; }
        public bool? HasExclamationMark { get; }
        public bool? HasCode { get; }
        public double? CoherenceCoefficient { get; }

        public Metrics(string content, int lineEnd, string type, CommentLocationStore commentLocationstore)
        {
            char[] delimiters = new char[] { ' ', '\t', '\r', '\n' };
            if (type == "single_line_comment")
            {
                string normalizedContent = WordTransform.RemoveSpecialCharacters(content);
                WordsCount = normalizedContent.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            if (commentLocationstore.CommentLocations.ContainsKey(lineEnd))
            {
                CommentLocation = commentLocationstore.CommentLocations[lineEnd][0];
                MethodName = commentLocationstore.CommentLocations[lineEnd][1];
            }

            HasNothing = new Regex("nothing", RegexOptions.IgnoreCase).IsMatch(content);
            HasQuestionMark = new Regex(@"\?($|\W)").IsMatch(content);
            HasExclamationMark = new Regex("!").IsMatch(content);
            HasCode = CodeDetector.HasCode(content);

            if (MethodName != null && CommentLocation == "method_description")
            {
                CoherenceCoefficient = RoslynPlay.CoherenceCoefficient.Compute(content, MethodName);
            }
        }

    }
}
