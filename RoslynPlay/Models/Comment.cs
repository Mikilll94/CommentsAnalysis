using System;

namespace RoslynPlay
{
    class Comment
    {
        public string FileName { get; set; }
        public int LineNumber { get; set; }
        public string Type { get; set; }

        public string Content { get; }
        public int LineEnd { get; }
        public string CommentLocation { get; } = "unknown";
        public int WordsCount { get; }

        public Comment(string content, int lineEnd)
        {
            Content = content;
            char[] delimiters = new char[] { ' ', '\r', '\n' };
            WordsCount = Content.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;

            LineEnd = lineEnd;

            if (CommentLocationStore.CommentLocations.ContainsKey(LineEnd))
            {
                CommentLocation = CommentLocationStore.CommentLocations[LineEnd];
            }
        }
    }
}
