using System.Collections.Generic;

namespace RoslynPlay
{
    public class CommentLocationStore
    {
        public Dictionary<int, string[]> CommentLocations { get; } = new Dictionary<int, string[]>();

        public void AddCommentLocation(int line, string location, string methodName)
        {
            CommentLocations.TryAdd(line, new string[] { location, methodName });
        }

        public void AddCommentLocation(int startLine, int endLine, string location, string methodName)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                CommentLocations.TryAdd(i, new string[] { location, methodName });
            }
        }
    }
}
