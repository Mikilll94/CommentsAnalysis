using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class DocComment : Comment
    {
        public DocComment(string content, int lineStart, int lineEnd, CommentLocationStore commentLocationstore)
        {
            Content = Regex.Replace(content, @"(\/\/\/)", "");
            Content = Regex.Replace(Content, "(<.*?>)", "");
            Type = "doc_comment";
            LineStart = lineStart;
            LineEnd = lineEnd;
            Statistics = new Statistics(Content, lineEnd, Type, commentLocationstore);
        }

        public override string GetLinesRange()
        {
            return $"{LineStart}-{LineEnd}";
        }
    }
}
