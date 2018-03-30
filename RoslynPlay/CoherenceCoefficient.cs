using System;

namespace RoslynPlay
{
    public static class CoherenceCoefficient
    {
        public static double Compute(string[] commentWords, string[] methodWords)
        {
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
