using ExcelDataReader;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var stream = File.Open("c:/Users/wasni/Desktop/Designite_GitExtensions.xls", FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    List<string> classes = new List<string>();

                    for (int i = 0; i < result.Tables["GitExtensions_AbsSMells"].Rows.Count; i++)
                    {
                        classes.Add(result.Tables["GitExtensions_AbsSMells"].Rows[i][2].ToString());
                    }
                    classes = classes.Distinct().ToList();

                    Console.WriteLine(String.Join("\n", classes));


                }

                //    ExcelWorksheet absSmellsWorksheet = codeSmellsExcelPackage.Workbook.Worksheets["GitExtensions_AbsSMells"];
                //int rowCount = absSmellsWorksheet.Dimension.End.Row;

                //List<string> classes = new List<string>();

                //for (int i = 1; i <= rowCount; i++)
                //{
                //    classes.Add(absSmellsWorksheet.Cells[i, 3].ToString());
                //}
                //Console.WriteLine(String.Join("\n", classes));
            }

            //string fileContent;
            //SyntaxTree tree;
            //SyntaxNode root;
            //CommentSyntaxWalker commentWalker;
            //MethodAndClassesSyntaxWalker methodWalker;
            //string projectName = "gitextensions";
            //string[] files = Directory.GetFiles($"c:/Users/wasni/Desktop/{projectName}", $"*.cs", SearchOption.AllDirectories);
            //var commentStore = new CommentStore();

            //Console.WriteLine("Reading files...");
            //ProgressBar progressBar = new ProgressBar(files.Length);

            //foreach (var file in files)
            //{
            //    fileContent = File.ReadAllText(file);
            //    string filePath = new Regex($@"{projectName}\\(.*)$").Match(file).Groups[1].ToString();
            //    tree = CSharpSyntaxTree.ParseText(fileContent);
            //    root = tree.GetRoot();
            //    var locationStore = new LocationStore();
            //    methodWalker = new MethodAndClassesSyntaxWalker(filePath, locationStore);
            //    methodWalker.Visit(root);
            //    commentWalker = new CommentSyntaxWalker(filePath, locationStore, commentStore);
            //    commentWalker.Visit(root);

            //    progressBar.UpdateAndDisplay();
            //}

            //Console.WriteLine("\nCreating excel file...");

            //ExcelWriter excelWriter = new ExcelWriter("comments.xlsx");
            //excelWriter.Write(commentStore);

            //Console.WriteLine("Finished");
            Console.ReadKey();
        }
    }
}
