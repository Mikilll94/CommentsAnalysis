using System.Collections.Generic;

namespace RoslynPlay
{
    public class LocationStore
    {
        public Dictionary<int, dynamic> MethodLocations { get; } = new Dictionary<int, dynamic>();
        public Dictionary<int, dynamic> ClassLocations { get; } = new Dictionary<int, dynamic>();

        public void AddClassLocation(int line, string location, Class @class)
        {
            ClassLocations.TryAdd(line, new { Location = location, Class = @class });
        }

        public void AddClassLocation(int startLine, int endLine, string location, Class @class)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                ClassLocations.TryAdd(i, new { Location = location, Class = @class });
            }
        }

        public void AddMethodLocation(int line, string location, string methodName)
        {
            MethodLocations.TryAdd(line, new { Location = location, Method = methodName });
        }

        public void AddMethodLocation(int startLine, int endLine, string location, string methodName)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                MethodLocations.TryAdd(i, new { Location = location, Method = methodName });
            }
        }
    }
}
