namespace RoslynPlay
{
    public abstract class Comment
    {
        public string FileName { get; set; }
        public string Content { get; set; }
        public int LineStart { get; set; }
        public int LineEnd { get; set; }
        public string Type { get; set; }
        public Metrics Metrics { get; set; }
        public EvaluationBad Evaluation { get; set; }

        public Comment()
        {
        }

        public abstract string GetLinesRange();
    }
}
