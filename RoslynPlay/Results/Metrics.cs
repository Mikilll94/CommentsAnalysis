using System;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public class Metrics
    {
        public LocationRelativeToMethod LocationRelativeToMethod { get; }
        public LocationRelativeToClass LocationRelativeToClass { get; }
        public string MethodName { get; }
        public string ClassName { get; }
        public bool? IsClassSmelly { get; }
        public bool? IsClassSmellyAbstraction { get; }
        public bool? IsClassSmellyEncapsulation { get; }
        public bool? IsClassSmellyModularization { get; }
        public bool? IsClassSmellyHierarchy { get; }
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

            if (locationstore.LocationsRelativeToMethod.ContainsKey(lineEnd))
            {
                LocationRelativeToMethod = locationstore.LocationsRelativeToMethod[lineEnd].Location;
                MethodName = locationstore.LocationsRelativeToMethod[lineEnd].Method;

            }
            if (locationstore.LocationsRelativeToClass.ContainsKey(lineEnd))
            {
                LocationRelativeToClass = locationstore.LocationsRelativeToClass[lineEnd].Location;
                ClassName = locationstore.LocationsRelativeToClass[lineEnd].Class.Name;
                IsClassSmelly = locationstore.LocationsRelativeToClass[lineEnd].Class.IsSmelly;
                IsClassSmellyAbstraction = locationstore.LocationsRelativeToClass[lineEnd].Class.IsSmellyAbstraction;
                IsClassSmellyEncapsulation = locationstore.LocationsRelativeToClass[lineEnd].Class.IsSmellyEncapsulation;
                IsClassSmellyModularization = locationstore.LocationsRelativeToClass[lineEnd].Class.IsSmellyModularization;
                IsClassSmellyHierarchy = locationstore.LocationsRelativeToClass[lineEnd].Class.IsSmellyHierarchy;
            }

            HasNothing = new Regex("nothing", RegexOptions.IgnoreCase).IsMatch(content);
            HasQuestionMark = new Regex(@"\?($|\W)").IsMatch(content);
            HasExclamationMark = new Regex("!").IsMatch(content);
            HasCode = CodeDetector.HasCode(content);

            if (MethodName != null && LocationRelativeToMethod == LocationRelativeToMethod.MethodDescription)
            {
                CoherenceCoefficient = RoslynPlay.CoherenceCoefficient.Compute(content, MethodName);
            }
        }

    }
}
