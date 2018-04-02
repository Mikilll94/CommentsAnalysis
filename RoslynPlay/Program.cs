using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileContent;
            SyntaxTree tree;
            SyntaxNode root;
            CommentSyntaxWalker commentWalker;
            MethodSyntaxWalker methodWalker;
            string projectName = "gitextensions";
            string[] files = Directory.GetFiles($"c:/Users/wasni/Desktop/{projectName}", $"*.cs", SearchOption.AllDirectories);
            var commentStore = new CommentStore();

            Console.WriteLine("Reading files...");
            ProgressBar progressBar = new ProgressBar(files.Length);

            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                string filePath = new Regex($@"{projectName}\\(.*)$").Match(file).Groups[1].ToString();
                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                var commentLocationStore = new CommentLocationStore();
                methodWalker = new MethodSyntaxWalker(filePath, commentLocationStore);
                methodWalker.Visit(root);
                commentWalker = new CommentSyntaxWalker(filePath, commentLocationStore, commentStore);
                commentWalker.Visit(root);

                progressBar.UpdateAndDisplay();
            }

            Console.WriteLine("\nCreating excel file...");

            ExcelWriter excelWriter = new ExcelWriter("comments.xlsx");
            excelWriter.Write(commentStore);

            Console.WriteLine("Finished");
            Console.ReadKey();
        }
    }
}
