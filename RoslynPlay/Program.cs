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

        static string projectPath = "c:/Users/wasni/Desktop/comments_analysis_data/EntityFrameworkCore/EntityFrameworkCore";
        static string smellsExcelPath = "c:/Users/wasni/Desktop/comments_analysis_data/EntityFrameworkCore/Designite_EFCore.xls";
        static string smellsSheetPrefix = "EFCore";

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
                methodWalker = new MethodsAndClassesWalker(filePath, locationStore);
                methodWalker.Visit(root);
                commentWalker = new CommentsWalker(filePath, locationStore, commentStore);
                commentWalker.Visit(root);

                progressBar.UpdateAndDisplay();
            }

            Console.WriteLine("\nCreating excel file...");

            ExcelWriter excelWriter = new ExcelWriter($"{smellsSheetPrefix}_comments.xlsx");
            excelWriter.Write(commentStore);

            Console.WriteLine("Finished");
            Console.ReadKey();
        }

        private static string TransformSingleLineComments(string fileContent)
        {
            string[] lines = fileContent.Split(Environment.NewLine);
            Regex singleLineCommentRegex = new Regex(@"^\/\/");
            for (int i = 0; i < lines.Length; i++)
            {
                if (singleLineCommentRegex.IsMatch(lines[i]))
                {
                    string line = lines[i];
                    int startIndex = i;
                    i++;
                    while(singleLineCommentRegex.IsMatch(lines[i]))
                    {
                        line += $" {lines[i].Replace(@"//", String.Empty)}";
                        lines[i] = String.Empty;
                        i++;
                    }
                    lines[startIndex] = line;
                }
            }
            fileContent = String.Join(Environment.NewLine, lines);
            return fileContent;
        }
    }
}
