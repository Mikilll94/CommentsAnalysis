using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public static class CoherenceCoefficient
    {
        public static double Compute(string commentContent, string methodName)
        {
            string[] methodWords = Regex.Replace(methodName, "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))", "$1 ").Split(" ");
            methodWords = methodWords.Where(word => word.Length > 2).ToArray();

            commentContent = Regex.Replace(commentContent, "[.,]", "");
            string[] commentWords = commentContent.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            commentWords = commentWords.Where(word => word.Length > 2).ToArray();
            commentWords = commentWords.Distinct().ToArray();

            if (commentWords.Length == 0)
            {
                return 0;
            }

            int matchedWords = 0;
            foreach (var commentWord in commentWords)
            {
                if (Array.Exists(methodWords, methodWord => methodWord.ToLower() == commentWord.ToLower()))
                {
                    matchedWords++;
                }
            }
            return (double)matchedWords/commentWords.Length;
        }
    }
}
