using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class DocComment : Comment
    {
        public DocComment(string content, int lineStart, int lineEnd, LocationStore commentLocationstore)
        {
            Content = Regex.Replace(content, @"(\/\/\/)", "");
            Content = Regex.Replace(Content, "(<.*?>)", "");
            Type = "doc_comment";
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
