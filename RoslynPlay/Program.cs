using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class Program
    {
        //static string projectPath = "c:/Users/wasni/Desktop/comments_analysis_data/gitextensions/gitextensions-master";
        //static string smellsExcelPath = "c:/Users/wasni/Desktop/comments_analysis_data/gitextensions/Designite_GitExtensions.xls";
        //static string smellsSheetPrefix = "GitExtensions";

        //static string projectPath = "c:/Users/wasni/Desktop/comments_analysis_data/EntityFrameworkCore/EntityFrameworkCore";
        //static string smellsExcelPath = "c:/Users/wasni/Desktop/comments_analysis_data/EntityFrameworkCore/Designite_EFCore.xls";
        //static string smellsSheetPrefix = "EFCore";

        static string projectPath = "c:/Users/wasni/Desktop/comments_analysis_data/ScreenToGif/ScreenToGif-master";
        static string smellsExcelPath = "c:/Users/wasni/Desktop/comments_analysis_data/ScreenToGif/Designite_GifRecorder.xls";
        static string smellsSheetPrefix = "GifRecorder";

        static void Main(string[] args)
        {
            SmellyClasses smellyClasses = new SmellyClasses(smellsExcelPath, smellsSheetPrefix);

            string fileContent;
            SyntaxTree tree;
            SyntaxNode root;
            CommentsWalker commentWalker;
            MethodsAndClassesWalker methodWalker;
            string[] files = Directory.GetFiles(projectPath, $"*.cs", SearchOption.AllDirectories);
            var commentStore = new CommentStore();
            var classStore = new ClassStore();

            Console.WriteLine("Reading files...");
            ProgressBar progressBar = new ProgressBar(files.Length);


            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                fileContent = TransformSingleLineComments(fileContent);

                string filePath = new Regex($@"{projectPath}\\(.*)$").Match(file).Groups[1].ToString();
                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                var locationStore = new LocationStore();
                methodWalker = new MethodsAndClassesWalker(filePath, locationStore, classStore);
                methodWalker.Visit(root);
                commentWalker = new CommentsWalker(filePath, locationStore, commentStore);
                commentWalker.Visit(root);

                progressBar.UpdateAndDisplay();
            }

            Console.WriteLine("\nCreating excel file...");

            ExcelWriter excelWriter = new ExcelWriter($"{smellsSheetPrefix}_comments.xlsx");
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
