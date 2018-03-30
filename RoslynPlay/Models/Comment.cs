namespace RoslynPlay
{
    public abstract class Comment
    {
        public string FileName { get; set; }
        public string Content { get; set; }
        public int LineStart { get; set; }
        public int LineEnd { get; set; }
        public string Type { get; set; }
        public Statistics Statistics { get; set; }

        public Comment()
        {
        }

        public abstract string GetLinesRange();
    }
}
