using System.Collections.Generic;

namespace RoslynPlay
{
    public enum LocationRelativeToMethod
    {
        MethodDescription,
        MethodStart,
        MethodInner,
        MethodEnd,
    };

    public enum LocationRelativeToClass
    {
        ClassDescription,
        ClassStart,
        ClassInner,
        ClassEnd,
    };

    public class LocationRelativeToMethodInfo
    {
        public LocationRelativeToMethod Location { get; set; }
        public string Method { get; set; }
    }

    public class LocationRelativeToClassInfo
    {
        public LocationRelativeToClass Location { get; set; }
        public Class @Class { get; set; }
    }

    public class LocationStore
    {
        public Dictionary<int, LocationRelativeToMethodInfo> LocationsRelativeToMethod { get; }
            = new Dictionary<int, LocationRelativeToMethodInfo>();
        public Dictionary<int, LocationRelativeToClassInfo> LocationsRelativeToClass { get; }
            = new Dictionary<int, LocationRelativeToClassInfo>();

        public void AddLocationRelativeToMethod(int line, LocationRelativeToMethod location, string methodName)
        {
            LocationsRelativeToMethod.TryAdd(line, new LocationRelativeToMethodInfo() { Location = location, Method = methodName });
        }

        public void AddLocationRelativeToMethod(int startLine, int endLine, LocationRelativeToMethod location, string methodName)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                LocationsRelativeToMethod.TryAdd(i, new LocationRelativeToMethodInfo() { Location = location, Method = methodName });
            }
        }

        public void AddLocationRelativeToClass(int line, LocationRelativeToClass location, Class @class)
        {
            LocationsRelativeToClass.TryAdd(line, new LocationRelativeToClassInfo() { Location = location, Class = @class });
        }

        public void AddLocationRelativeToClass(int startLine, int endLine, LocationRelativeToClass location, Class @class)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                LocationsRelativeToClass.TryAdd(i, new LocationRelativeToClassInfo() { Location = location, Class = @class });
            }
        }
    }
}
