using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using System.Text;

namespace RoslynPlay
{
    class Program
    {
        static void FormatCells(ExcelRange cells, bool condition)
        {
            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            if (condition)
            {
                cells.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            }
            else
            {
                cells.Style.Fill.BackgroundColor.SetColor(Color.IndianRed);
            }
        }

        static void Main(string[] args)
        {
            string fileContent;
            SyntaxTree tree;
            SyntaxNode root;
            CommentSyntaxWalker commentWalker;
            MethodSyntaxWalker methodWalker;
            string[] files = Directory.GetFiles($"c:/Users/wasni/Desktop/gitextensions-master", $"*.cs", SearchOption.AllDirectories);
            var commentStore = new CommentStore();
            foreach (var file in files)
            {
                fileContent = File.ReadAllText(file);
                tree = CSharpSyntaxTree.ParseText(fileContent);
                root = tree.GetRoot();
                var commentLocationStore = new CommentLocationStore();
                methodWalker = new MethodSyntaxWalker(file, commentLocationStore);
                methodWalker.Visit(root);
                commentWalker = new CommentSyntaxWalker(file, commentLocationStore, commentStore);
                commentWalker.Visit(root);
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            FileInfo xlsxFile = new FileInfo(@"comments.xlsx");
            using (ExcelPackage package = new ExcelPackage(xlsxFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Worksheet");
                worksheet.Cells[1, 1].Value = "File";
                worksheet.Cells[1, 2].Value = "Line";
                worksheet.Cells[1, 3].Value = "Type";
                worksheet.Cells[1, 4].Value = "Words count";
                worksheet.Cells[1, 5].Value = "Has \"nothing\"";
                worksheet.Cells[1, 6].Value = "Has !";
                worksheet.Cells[1, 7].Value = "Has ?";
                worksheet.Cells[1, 8].Value = "Location";
                worksheet.Cells[1, 9].Value = "Method name";
                worksheet.Cells[1, 10].Value = "Comment";

                int rowNumber = 2;

                foreach (var comment in commentStore.Comments)
                {
                    ExcelRange wordsCountCell = worksheet.Cells[rowNumber, 4];
                    ExcelRange hasNothingCell = worksheet.Cells[rowNumber, 5];
                    ExcelRange hasExclamationMarkCell = worksheet.Cells[rowNumber, 6];
                    ExcelRange hasQuestionMarkCell = worksheet.Cells[rowNumber, 7];

                    worksheet.Cells[rowNumber, 1].Value = comment.FileName;
                    worksheet.Cells[rowNumber, 2].Value
                        = comment.LineEnd != -1 ? $"{comment.LineNumber}-{comment.LineEnd}" : comment.LineNumber.ToString();
                    worksheet.Cells[rowNumber, 3].Value = comment.Type;
                    wordsCountCell.Value = comment.WordsCount;
                    hasNothingCell.Value = comment.HasNothing;
                    hasExclamationMarkCell.Value = comment.HasExclamationMark;
                    hasQuestionMarkCell.Value = comment.HasQuestionMark;
                    worksheet.Cells[rowNumber, 8].Value = comment.CommentLocation;
                    worksheet.Cells[rowNumber, 9].Value = comment.MethodName;
                    worksheet.Cells[rowNumber, 10].Value = comment.Content;

                    if (wordsCountCell.Value != null && wordsCountCell.Value.ToString() != "0")
                    {
                        FormatCells(wordsCountCell, int.Parse(wordsCountCell.Value.ToString()) > 2);
                    }
                    FormatCells(hasNothingCell, bool.Parse(hasNothingCell.Value.ToString()));
                    FormatCells(hasExclamationMarkCell, bool.Parse(hasExclamationMarkCell.Value.ToString()));
                    FormatCells(hasQuestionMarkCell, bool.Parse(hasQuestionMarkCell.Value.ToString()));

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

                package.Save();
            }
            Console.WriteLine("Finished");
            Console.ReadKey();
        }
    }
}
