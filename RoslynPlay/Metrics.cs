using System;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public class Metrics
    {
        public string LocationMethod { get; }
        public string LocationClass { get; }
        public string MethodName { get; }
        public string ClassName { get; }
        public bool? IsClassSmelly { get; }
        public int? WordsCount { get; }
        public bool? HasNothing { get; }
        public bool? HasQuestionMark { get; }
        public bool? HasExclamationMark { get; }
        public bool? HasCode { get; }
        public double? CoherenceCoefficient { get; }

        public Metrics(string content, int lineEnd, string type, LocationStore locationstore)
        {
            char[] delimiters = new char[] { ' ', '\t', '\r', '\n' };
            if (type == "single_line_comment")
            {
                string normalizedContent = WordTransform.RemoveSpecialCharacters(content);
                WordsCount = normalizedContent.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            if (locationstore.MethodLocations.ContainsKey(lineEnd))
            {
                LocationMethod = locationstore.MethodLocations[lineEnd][0];
                MethodName = locationstore.MethodLocations[lineEnd][1];

            }
            if (locationstore.ClassLocations.ContainsKey(lineEnd))
            {
                LocationClass = locationstore.ClassLocations[lineEnd].Location;
                ClassName = locationstore.ClassLocations[lineEnd].Name;
                IsClassSmelly = locationstore.ClassLocations[lineEnd].IsSmelly;
            }

            HasNothing = new Regex("nothing", RegexOptions.IgnoreCase).IsMatch(content);
            HasQuestionMark = new Regex(@"\?($|\W)").IsMatch(content);
            HasExclamationMark = new Regex("!").IsMatch(content);
            HasCode = CodeDetector.HasCode(content);

            if (MethodName != null && LocationMethod == "method_description")
            {
                CoherenceCoefficient = RoslynPlay.CoherenceCoefficient.Compute(content, MethodName);
            }
        }

    }
}
