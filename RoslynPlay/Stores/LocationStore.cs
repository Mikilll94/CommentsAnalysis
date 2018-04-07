using System.Collections.Generic;

namespace RoslynPlay
{
    public class LocationStore
    {
        public Dictionary<int, string[]> MethodLocations { get; } = new Dictionary<int, string[]>();
        public Dictionary<int, Class> ClassLocations { get; } = new Dictionary<int, Class>();

        public void AddClassLocation(int line, string location, string className)
        {
            ClassLocations.TryAdd(line, 
                new Class
                {
                    Location = location,
                    Name = className,
                    IsSmelly = SmellyClasses.All.Contains(className),
                    IsSmellyAbstraction = SmellyClasses.Abstraction.Contains(className),
                    IsSmellyEncapsulation = SmellyClasses.Encapsulation.Contains(className),
                    IsSmellyModularization = SmellyClasses.Modularization.Contains(className),
                    IsSmellyHierarchy = SmellyClasses.Hierarchy.Contains(className),
                });
        }

        public void AddClassLocation(int startLine, int endLine, string location, string className)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                ClassLocations.TryAdd(i,
                    new Class
                    {
                        Location = location,
                        Name = className,
                        IsSmelly = SmellyClasses.All.Contains(className),
                        IsSmellyAbstraction = SmellyClasses.Abstraction.Contains(className),
                        IsSmellyEncapsulation = SmellyClasses.Encapsulation.Contains(className),
                        IsSmellyModularization = SmellyClasses.Modularization.Contains(className),
                        IsSmellyHierarchy = SmellyClasses.Hierarchy.Contains(className),
                    });
            }
        }

        public void AddMethodLocation(int line, string location, string methodName)
        {
            MethodLocations.TryAdd(line, new string[] { location, methodName });
        }

        public void AddMethodLocation(int startLine, int endLine, string location, string methodName)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                MethodLocations.TryAdd(i, new string[] { location, methodName });
            }
        }
    }
}
