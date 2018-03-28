using System.Collections.Generic;

namespace RoslynPlay
{
    public class CommentLocationStore
    {
        public Dictionary<int, string> CommentLocations { get; } = new Dictionary<int, string>();

        public void AddCommentLocation(int line, string location)
        {
            CommentLocations.TryAdd(line, location);
        }

        public void AddCommentLocation(int startLine, int endLine, string location)
        {
            for (int i = startLine; i <= endLine; i++)
            {
                CommentLocations.TryAdd(i, location);
            }
        }
    }
}
