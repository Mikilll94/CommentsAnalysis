namespace RoslynPlay
{
    class SingleLineComment : Comment
    {
        public SingleLineComment(string content, int line, LocationStore commentLocationstore)
        {
            Content = content.Substring(content.IndexOf("//") + 2);
            Type = "single_line_comment";
            LineStart = line;
            LineEnd = line;
            Metrics = new Metrics(Content, line, Type, commentLocationstore);
            Evaluation = new EvaluationBad(Metrics);
        }

        public override string GetLinesRange()
        {
            return LineStart.ToString();
        }
    }
}
