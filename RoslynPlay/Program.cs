using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class Program
    {
        static string projectPath = "c:/Users/wasni/Desktop/comments_analysis_data/gitextensions/gitextensions";
        static string smellsExcelPath = "c:/Users/wasni/Desktop/comments_analysis_data/gitextensions/Designite_GitExtensions.xls";
        static string smellsSheetPrefix = "GitExtensions";

        static void Main(string[] args)
        {
            SmellyClasses smellyClasses = new SmellyClasses(smellsExcelPath, smellsSheetPrefix);

            string fileContent;
            SyntaxTree tree;
            SyntaxNode root;
            CommentSyntaxWalker commentWalker;
            MethodAndClassesSyntaxWalker methodWalker;
            string[] files = Directory.GetFiles(projectPath, $"*.cs", SearchOption.AllDirectories);
            var commentStore = new CommentStore();

            Console.WriteLine("Reading files...");
            ProgressBar progressBar = new ProgressBar(files.Length);

            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                string filePath = new Regex($@"{projectPath}\\(.*)$").Match(file).Groups[1].ToString();
                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                var locationStore = new LocationStore();
                methodWalker = new MethodAndClassesSyntaxWalker(filePath, locationStore);
                methodWalker.Visit(root);
                commentWalker = new CommentSyntaxWalker(filePath, locationStore, commentStore);
                commentWalker.Visit(root);

                progressBar.UpdateAndDisplay();
            }

            Console.WriteLine("\nCreating excel file...");

            ExcelWriter excelWriter = new ExcelWriter($"{smellsSheetPrefix}_comments.xlsx");
            excelWriter.Write(commentStore);

            Console.WriteLine("Finished");
            Console.ReadKey();
        }
    }
}
