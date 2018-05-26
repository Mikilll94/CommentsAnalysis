namespace CommentsAnalysis.Models
{
    public enum SmellType
    {
        Abstraction,
        Encapsulation,
        Modularization,
        Hierarchy
    }

    public class Smell
    {
        public SmellType Type { get; set; }
        public string ClassName { get; set; }
        public string ClassNamespace { get; set; }
    }
}
