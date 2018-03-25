using System;

namespace RoslynPlay
{
    class Comment
    {
        public string FileName { get; set; }
        public int LineNumber { get; set; }
        public string Type { get; set; }

        public int LineEnd { get; set; } = -1;

        public string Content { get; }
        public int WordsCount { get; }


        public Comment(string content)
        {
            Content = content;
            char[] delimiters = new char[] { ' ', '\r', '\n' };
            WordsCount = Content.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
        }
    }
}
