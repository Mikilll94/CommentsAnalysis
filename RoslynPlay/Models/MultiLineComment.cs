using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class MultiLineComment : Comment
    {
        public MultiLineComment(string content, int lineStart, int lineEnd, LocationStore commentLocationstore)
        {
            Content = new Regex(@"\/\*(.*)\*\/", RegexOptions.Singleline).Match(content).Groups[1].ToString();
            Type = CommentType.MultiLine;
            LineStart = lineStart;
            LineEnd = lineEnd;
            Metrics = new Metrics(Content, lineEnd, Type, commentLocationstore);
            Evaluation = new EvaluationBad(Metrics);
        }

        public override string GetLinesRange()
        {
            return $"{LineStart}-{LineEnd}";
        }
    }
}
