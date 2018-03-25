using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using OfficeOpenXml;
using System;
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
            SyntaxWalker walker;
            string[] files = Directory.GetFiles($"c:/Users/wasni/Desktop/gitextensions-master", $"*.cs", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                walker = new SyntaxWalker(file);
                walker.Visit(root);
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
                worksheet.Cells[1, 5].Value = "Words count";

                int rowNumber = 2;

                foreach (var comment in CommentsStore.Comments)
                {
                    worksheet.Cells[rowNumber, 1].Value = comment.FileName;
                    worksheet.Cells[rowNumber, 2].Value
                        = comment.LineEnd != -1 ? $"{comment.LineNumber}-{comment.LineEnd}" : comment.LineNumber.ToString();
                    worksheet.Cells[rowNumber, 3].Value = comment.Content;
                    worksheet.Cells[rowNumber, 4].Value = comment.Type;
                    worksheet.Cells[rowNumber, 5].Value = comment.WordsCount;
                    rowNumber++;
                }
                package.Save();
            }
            Console.WriteLine("Finished");
            Console.ReadKey();
        }
    }
}
