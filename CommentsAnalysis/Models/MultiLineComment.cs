using System.Text.RegularExpressions;

namespace CommentsAnalysis
{
    class MultiLineComment : Comment
    {
        public MultiLineComment(string content, int lineStart, int lineEnd)
        {
            Content = new Regex(@"\/\*(.*)\*\/", RegexOptions.Singleline).Match(content).Groups[1].ToString();
            Type = CommentType.MultiLine;
            LineStart = lineStart;
            LineEnd = lineEnd;
        }

        public override string GetLinesRange()
        {
            return $"{LineStart}-{LineEnd}";
        }
    }
}
