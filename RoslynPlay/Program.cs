using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

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
            string[] files = Directory.GetFiles($"c:/Users/wasni/Desktop/gitextensions-master", $"*.cs", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                CommentLocationStore.CommentLocations.Clear();
                methodWalker = new MethodSyntaxWalker(file);
                methodWalker.Visit(root);
                commentWalker = new CommentSyntaxWalker(file);
                commentWalker.Visit(root);
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            FileInfo xlsxFile = new FileInfo(@"comments.xlsx");
            using (ExcelPackage package = new ExcelPackage(xlsxFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Worksheet");
                worksheet.Cells[1, 1].Value = "File";
                worksheet.Cells[1, 2].Value = "Line";
                worksheet.Cells[1, 3].Value = "Comment";
                worksheet.Cells[1, 4].Value = "Type";
                worksheet.Cells[1, 5].Value = "Location";
                worksheet.Cells[1, 6].Value = "Words count";

                int rowNumber = 2;

                foreach (var comment in CommentStore.Comments)
                {
                    worksheet.Cells[rowNumber, 1].Value = comment.FileName;
                    worksheet.Cells[rowNumber, 2].Value
                        = comment.LineEnd != -1 ? $"{comment.LineNumber}-{comment.LineEnd}" : comment.LineNumber.ToString();
                    worksheet.Cells[rowNumber, 3].Value = comment.Content;
                    worksheet.Cells[rowNumber, 4].Value = comment.Type;
                    worksheet.Cells[rowNumber, 5].Value = comment.CommentLocation;
                    worksheet.Cells[rowNumber, 6].Value = comment.WordsCount;
                    rowNumber++;
                }

                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Methods");
                worksheet2.Cells[1, 1].Value = "Name";
                worksheet2.Cells[1, 2].Value = "File";
                worksheet2.Cells[1, 3].Value = "Line";

                rowNumber = 2;

                foreach (var method in MethodStore.Methods)
                {
                    worksheet2.Cells[rowNumber, 1].Value = method.Name;
                    worksheet2.Cells[rowNumber, 2].Value = method.FileName;
                    worksheet2.Cells[rowNumber, 3].Value = $"{method.LineNumber}-{method.LineEnd}";
                    rowNumber++;
                }

                ExcelWorksheet worksheet3 = package.Workbook.Worksheets.Add("Locations");
                worksheet3.Cells[1, 1].Value = "Line";
                worksheet3.Cells[1, 2].Value = "Location";

                rowNumber = 2;

                foreach (var commentLocation in CommentLocationStore.CommentLocations)
                {
                    worksheet3.Cells[rowNumber, 1].Value = commentLocation.Key;
                    worksheet3.Cells[rowNumber, 2].Value = commentLocation.Value;
                    rowNumber++;
                }

                package.Save();
            }
            Console.WriteLine("Finished");
            Console.ReadKey();
        }
    }
}
