namespace RoslynPlay
{
    class SingleLineComment : Comment
    {
        public SingleLineComment(string content, int line, CommentLocationStore commentLocationstore)
        {
            Content = content.Substring(content.IndexOf("//") + 2);
            Type = "single_line_comment";
            LineStart = line;
            LineEnd = line;
            Statistics = new Statistics(Content, line, Type, commentLocationstore);
        }

        public override string GetLinesRange()
        {
            return LineStart.ToString();
        }
    }
}
