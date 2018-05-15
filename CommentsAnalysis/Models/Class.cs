namespace RoslynPlay
{
    public class Class
    {
        public string Name { get; set; }
        public string FileName { get; set; }
        public int SmellsCount { get; set; }
        public int AbstractionSmellsCount { get; set; }
        public int EncapsulationSmellsCount { get; set; }
        public int ModularizationSmellsCount { get; set; }
        public int HierarchySmellsCount { get; set; }

        public bool IsSmelly => SmellsCount > 0;
        public bool IsSmellyAbstraction => AbstractionSmellsCount > 0;
        public bool IsSmellyEncapsulation => EncapsulationSmellsCount > 0;
        public bool IsSmellyModularization => ModularizationSmellsCount > 0;
        public bool IsSmellyHierarchy => HierarchySmellsCount > 0;
    }
}
