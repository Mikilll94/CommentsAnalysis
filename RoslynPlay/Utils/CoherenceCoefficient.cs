using Iveonik.Stemmers;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public static class CoherenceCoefficient
    {
        private static EnglishStemmer stemmer = new EnglishStemmer();

        public static string[] PreprocessWordsArray(string [] wordsArray)
        {
            wordsArray = wordsArray.Where(word => word.Length > 2).ToArray();
            wordsArray = wordsArray.Select(word => stemmer.Stem(word)).ToArray();
            return wordsArray;
        }

        public static double Compute(string commentContent, string methodName)
        {
            string[] methodWords = Regex.Replace(methodName, "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))", "$1 ").Split(" ");
            methodWords = PreprocessWordsArray(methodWords);

            commentContent = Regex.Replace(commentContent, "[.,]", "");
            string[] commentWords = commentContent.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            commentWords = PreprocessWordsArray(commentWords);

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
