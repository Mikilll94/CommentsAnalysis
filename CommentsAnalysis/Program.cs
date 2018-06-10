using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace CommentsAnalysis
{
    class Program
    {
        static string folderName = "gitextensions";
        static string designiteFileName = "Designite_GitExtensions.xls";
        static string solutionName = "GitExtensions";

        //static string folderName = "EntityFrameworkCore";
        //static string designiteFileName = "Designite_EFCore.xls";
        //static string solutionName = "EFCore";

        //static string folderName = "ScreenToGif";
        //static string designiteFileName = "Designite_GifRecorder.xls";
        //static string solutionName = "GifRecorder";

        static void Main(string[] args)
        {
            SmellsStore.Initialize($@"../../../../DesigniteResults/{designiteFileName}", solutionName);

            string fileContent;
            SyntaxTree tree;
            SyntaxNode root;
            CommentsWalker commentWalker;
            MethodsAndClassesWalker methodWalker;
            string[] files = Directory.GetFiles($@"../../../../Projects/{folderName}", $"*.cs", SearchOption.AllDirectories);
            var commentStore = new CommentStore();
            var classStore = new ClassStore();

            Console.WriteLine("Reading files...");
            ProgressBar progressBar = new ProgressBar(files.Length);

            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                fileContent = TransformSingleLineComments(fileContent);

                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                var locationStore = new LocationStore();
                string filePath = new Regex($@"{folderName}\\(.*)").Match(file).Groups[1].ToString();

                methodWalker = new MethodsAndClassesWalker(filePath, locationStore, classStore);
                methodWalker.Visit(root);
                commentWalker = new CommentsWalker(filePath, locationStore, commentStore, classStore);
                commentWalker.Visit(root);

                progressBar.UpdateAndDisplay();
            }

            Console.WriteLine("\nCreating excel file...");

            ExcelWriter excelWriter = new ExcelWriter($"{solutionName}_comments.xlsx");
            excelWriter.Write(commentStore, classStore);

            Console.WriteLine("Finished");
            Console.ReadKey();
        }

        private static string TransformSingleLineComments(string fileContent)
        {
            fileContent = fileContent.Replace("\r\n", "\n");
            string[] lines = fileContent.Split("\n");
            Regex singleLineCommentRegex = new Regex(@"^[\s]*\/\/");


            for (int i = 0; i < lines.Length; i++)
            {
                if (StartsWithSingleLineComment(lines[i]))
                {
                    string line = lines[i];
                    int startIndex = i;
                    i++;
                    while (StartsWithSingleLineComment(lines[i]))
                    {
                        line += $" {lines[i].Replace(@"//", String.Empty)}";
                        lines[i] = String.Empty;
                        i++;
                    }
                    lines[startIndex] = line;
                }
            }
            fileContent = String.Join("\n", lines);
            return fileContent;
        }

        private static bool StartsWithSingleLineComment(string line)
        {
            return line.TrimStart().StartsWith(@"//") && !line.TrimStart().StartsWith(@"///");
        }
    }
}
