using System.Collections.Generic;

namespace RoslynPlay
{
    static class CommentLocationStore
    {
        public static Dictionary<int, string> CommentLocations { get; set; } = new Dictionary<int, string>();

        public static void AddCommentLocation(int start, int end, string location)
        {
            for (int i = start; i <= end; i++)
            {
                CommentLocations.TryAdd(i, location);
            }
        }
    }
}
