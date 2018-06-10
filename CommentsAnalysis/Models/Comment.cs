using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CommentsAnalysis
{
    public enum CommentType
    {
        SingleLine,
        MultiLine,
        Doc
    }

    public abstract class Comment
    {
        public string FileName { get; set; }
        public string Content { get; set; }
        public int LineStart { get; set; }
        public int LineEnd { get; set; }
        public CommentType Type { get; set; }

        public LocationRelativeToMethod LocationRelativeToMethod { get; private set; }
        public LocationRelativeToClass LocationRelativeToClass { get; private set; }
        public string MethodName { get; private set; }
        public Class Class { get; private set; }

        public int? WordsCount { get; private set; }
        public bool? HasNothing { get; private set; }
        public bool? HasQuestionMark { get; private set; }
        public bool? HasExclamationMark { get; private set; }
        public bool? HasCode { get; private set; }
        public double? CoherenceCoefficient { get; private set; }

        public Comment()
        {

        }

        public void Initialize(LocationStore locationStore, ClassStore classStore)
        {
            char[] delimiters = new char[] { ' ', '\t', '\r', '\n' };
            if (Type == CommentType.SingleLine)
            {
                string normalizedContent = WordTransform.RemoveSpecialCharacters(Content);
                WordsCount = normalizedContent.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            if (locationStore.LocationsRelativeToMethod.ContainsKey(LineEnd))
            {
                LocationRelativeToMethod = locationStore.LocationsRelativeToMethod[LineEnd].Location;
                MethodName = locationStore.LocationsRelativeToMethod[LineEnd].Method;

            }
            if (locationStore.LocationsRelativeToClass.ContainsKey(LineEnd))
            {
                LocationRelativeToClass = locationStore.LocationsRelativeToClass[LineEnd].Location;
                Class @class = classStore.Classes.Single(c => c.Name == locationStore.LocationsRelativeToClass[LineEnd].Class.Name
                    && c.Namespace == locationStore.LocationsRelativeToClass[LineEnd].Class.Namespace);
                this.Class = @class;
            }

            HasNothing = new Regex("nothing", RegexOptions.IgnoreCase).IsMatch(Content);
            HasQuestionMark = new Regex(@"\?($|\W)").IsMatch(Content);
            HasExclamationMark = new Regex("!").IsMatch(Content);
            HasCode = CodeDetector.HasCode(Content);

            if (MethodName != null && LocationRelativeToMethod == LocationRelativeToMethod.MethodDescription)
            {
                CoherenceCoefficient = CommentsAnalysis.CoherenceCoefficient.Compute(Content, MethodName);
            }
        }

        public abstract string GetLinesRange();

        public bool IsBad()
        {
            if (HasNothing == true || HasQuestionMark == true || HasExclamationMark == true || HasCode == true)
            {
                return true;
            }
            if (CoherenceCoefficient != null && (CoherenceCoefficient == 0 || CoherenceCoefficient > 0.5))
            {
                return true;
            }
            if ((WordsCount != null || WordsCount != 0) && (WordsCount <= 2 || WordsCount > 30))
            {
                return true;
            }
            return false;
        }

        public bool? IsBadCoherenceCoefficient()
        {
            if (CoherenceCoefficient == null) return null;

            return CoherenceCoefficient == 0 || CoherenceCoefficient > 0.5;
        }

        public bool? IsBadWordsCount()
        {
            if (WordsCount == null || WordsCount == 0) return null;

            return WordsCount <= 2 || WordsCount > 30;
        }
    }
}
