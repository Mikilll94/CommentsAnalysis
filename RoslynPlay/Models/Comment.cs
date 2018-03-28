using System;

namespace RoslynPlay
{
    public class Comment
    {
        public string FileName { get; set; }
        public int LineNumber { get; set; }
        public string Type { get; set; }

        public string Content { get; }
        public int LineEnd { get; }
        public string CommentLocation { get; } = "Unknown";
        public string MethodName { get; } = "No method";
        public int WordsCount { get; }

        public Comment(string content, int lineEnd, CommentLocationStore commentLocationstore)
        {
            Content = content;
            LineEnd = lineEnd;

            char[] delimiters = new char[] { ' ', '\r', '\n' };
            WordsCount = Content.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;

            if (commentLocationstore.CommentLocations.ContainsKey(LineEnd))
            {
                CommentLocation = commentLocationstore.CommentLocations[LineEnd][0];
                MethodName = commentLocationstore.CommentLocations[LineEnd][1];
            }
        }


    }
}
