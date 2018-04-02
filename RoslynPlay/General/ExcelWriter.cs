using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Text;

namespace RoslynPlay
{
    class ExcelWriter
    {
        private FileInfo _file;

        public ExcelWriter(string file)
        {
            _file = new FileInfo(file);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void FormatCell(ExcelRange cells, bool? condition)
        {
            if (condition == null) return;

            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cells.Style.Fill.BackgroundColor.SetColor((bool)condition ? Color.IndianRed : Color.LightGreen);
        }

        public void Write(CommentStore commentStore)
        {
            using (ExcelPackage package = new ExcelPackage(_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Worksheet");
                worksheet.Cells[1, 1].Value = "File";
                worksheet.Cells[1, 2].Value = "Line";
                worksheet.Cells[1, 3].Value = "Type";
                worksheet.Cells[1, 4].Value = "Words count";
                worksheet.Cells[1, 5].Value = "Has \"nothing\"";
                worksheet.Cells[1, 6].Value = "Has !";
                worksheet.Cells[1, 7].Value = "Has ?";
                worksheet.Cells[1, 8].Value = "Has code";
                worksheet.Cells[1, 9].Value = "Coherence coefficient";
                worksheet.Cells[1, 10].Value = "Location";
                worksheet.Cells[1, 11].Value = "Method name";
                worksheet.Cells[1, 12].Value = "Comment";

                int rowNumber = 2;

                foreach (var comment in commentStore.Comments)
                {
                    ExcelRange wordsCountCell = worksheet.Cells[rowNumber, 4];
                    ExcelRange hasNothingCell = worksheet.Cells[rowNumber, 5];
                    ExcelRange hasExclamationMarkCell = worksheet.Cells[rowNumber, 6];
                    ExcelRange hasQuestionMarkCell = worksheet.Cells[rowNumber, 7];
                    ExcelRange hasCodeCell = worksheet.Cells[rowNumber, 8];
                    ExcelRange coherenceCoefficientCell = worksheet.Cells[rowNumber, 9];

                    worksheet.Cells[rowNumber, 1].Value = comment.FileName;
                    worksheet.Cells[rowNumber, 2].Value = comment.GetLinesRange();
                    worksheet.Cells[rowNumber, 3].Value = comment.Type;
                    wordsCountCell.Value = comment.Statistics.WordsCount;
                    hasNothingCell.Value = comment.Statistics.HasNothing;
                    hasExclamationMarkCell.Value = comment.Statistics.HasExclamationMark;
                    hasQuestionMarkCell.Value = comment.Statistics.HasQuestionMark;
                    hasCodeCell.Value = comment.Statistics.HasCode;
                    coherenceCoefficientCell.Value = comment.Statistics.CoherenceCoefficient;
                    worksheet.Cells[rowNumber, 10].Value = comment.Statistics.CommentLocation;
                    worksheet.Cells[rowNumber, 11].Value = comment.Statistics.MethodName;
                    worksheet.Cells[rowNumber, 12].Value = comment.Content;

                    FormatCell(wordsCountCell, comment.Statistics.WordsCountBad());
                    FormatCell(hasNothingCell, comment.Statistics.HasNothing);
                    FormatCell(hasExclamationMarkCell, comment.Statistics.HasExclamationMark);
                    FormatCell(hasQuestionMarkCell, comment.Statistics.HasQuestionMark);
                    FormatCell(hasCodeCell, comment.Statistics.HasCode);
                    FormatCell(coherenceCoefficientCell, comment.Statistics.CoherenceCoefficientBad());

                    rowNumber++;
                }

                for (int i = 2; i <= 10; i++)
                {
                    worksheet.Column(i).AutoFit();
                }

                package.Save();
            }
        }
    }
}
