using System.Collections.Generic;

namespace RoslynPlay
{
    public class LocationStore
    {
        public Dictionary<int, string[]> MethodLocations { get; } = new Dictionary<int, string[]>();
        public Dictionary<int, ClassLocation> ClassLocations { get; } = new Dictionary<int, ClassLocation>();

        public void AddClassLocation(int line, string location, Class @class)
        {
            ClassLocations.TryAdd(line, new ClassLocation
            {
                Location = location,
                Class = @class,
            });
        }

        public void AddClassLocation(int startLine, int endLine, string location, Class @class)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                ClassLocations.TryAdd(i, new ClassLocation
                {
                    Location = location,
                    Class = @class,
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
