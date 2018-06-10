namespace CommentsAnalysis
{
    class SingleLineComment : Comment
    {
        public SingleLineComment(string content, int line)
        {
            Content = content.Substring(content.IndexOf("//") + 2);
            Type = CommentType.SingleLine;
            LineStart = line;
            LineEnd = line;
        }

        public override string GetLinesRange()
        {
            return LineStart.ToString();
        }
    }
}
