using System.Text.RegularExpressions;

namespace CommentsAnalysis
{
    class DocComment : Comment
    {
        public DocComment(string content, int lineStart, int lineEnd)
        {
            Content = Regex.Replace(content, @"(\/\/\/)", "");
            Content = Regex.Replace(Content, "(<.*?>)", "");
            Type = CommentType.Doc;
            LineStart = lineStart;
            LineEnd = lineEnd;
        }

        public override string GetLinesRange()
        {
            return $"{LineStart}-{LineEnd}";
        }
    }
}
