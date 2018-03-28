using System;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public class Comment
    {
        public string FileName { get; set; }
        public int LineNumber { get; set; }

        public string Content { get; }
        public int LineEnd { get; }
        public string Type { get; }

        public string CommentLocation { get; } = "Unknown";
        public string MethodName { get; } = "No method";
        public int? WordsCount { get; }
        public bool? HasNothing { get; }
        public bool? HasQuestionMark { get; }
        public bool? HasExclamationMark { get; }

        public Comment(string content, int lineEnd, string type,
            CommentLocationStore commentLocationstore)
        {
            Content = content;
            LineEnd = lineEnd;
            Type = type;

            char[] delimiters = new char[] { ' ', '\r', '\n' };
            if (Type == "SingleLineCommentTrivia")
            {
                WordsCount = Content.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            if (commentLocationstore.CommentLocations.ContainsKey(LineEnd))
            {
                CommentLocation = commentLocationstore.CommentLocations[LineEnd][0];
                MethodName = commentLocationstore.CommentLocations[LineEnd][1];
            }

            HasNothing = new Regex("nothing", RegexOptions.IgnoreCase).IsMatch(Content);
            HasQuestionMark = new Regex(@"\?").IsMatch(Content);
            HasExclamationMark = new Regex("!").IsMatch(Content);
        }


    }
}
