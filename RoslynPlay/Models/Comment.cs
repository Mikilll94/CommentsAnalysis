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

        public Statistics Statistics { get; }

        public Comment(string content, int lineEnd, string type,
            CommentLocationStore commentLocationstore)
        {
            Content = content;
            LineEnd = lineEnd;
            Type = type;
            Statistics = new Statistics(content, lineEnd, type, commentLocationstore);
        }


    }
}
