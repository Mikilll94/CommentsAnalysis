namespace RoslynPlay
{
    public enum CommentType
    {
        SingleLine,
        MultiLine,
        Doc
    }

    public abstract class Comment
    {
        public string FileName { get; set; }
        public string Content { get; set; }
        public int LineStart { get; set; }
        public int LineEnd { get; set; }
        public CommentType Type { get; set; }
        public Metrics Metrics { get; set; }
        public EvaluationBad Evaluation { get; set; }

        public abstract string GetLinesRange();
    }
}
