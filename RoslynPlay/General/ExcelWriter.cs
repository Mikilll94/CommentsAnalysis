using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Linq;
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

        public void Write(CommentStore commentStore)
        {
            using (ExcelPackage package = new ExcelPackage(_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Comments");
                worksheet.View.FreezePanes(2, 1);
                worksheet.Cells[1, 1].Value = "File";
                worksheet.Cells[1, 2].Value = "Line";
                worksheet.Cells[1, 3].Value = "Type";
                worksheet.Cells[1, 4].Value = "Words count";
                worksheet.Cells[1, 5].Value = "Has \"nothing\"";
                worksheet.Cells[1, 6].Value = "Has !";
                worksheet.Cells[1, 7].Value = "Has ?";
                worksheet.Cells[1, 8].Value = "Has code";
                worksheet.Cells[1, 9].Value = "Coherence coefficient";
                worksheet.Cells[1, 10].Value = "Location method";
                worksheet.Cells[1, 11].Value = "Method name";
                worksheet.Cells[1, 12].Value = "Location class";
                worksheet.Cells[1, 13].Value = "Class name";
                worksheet.Cells[1, 14].Value = "Is class smelly?";
                worksheet.Cells[1, 15].Value = "Is bad?";
                worksheet.Cells[1, 16].Value = "Comment";

                int rowNumber = 2;

                foreach (var comment in commentStore.Comments)
                {
                    ExcelRowWriter excelRowWriter = new ExcelRowWriter(worksheet, rowNumber);

                    excelRowWriter.WriteCell(1, comment.FileName);
                    excelRowWriter.WriteCell(2, comment.GetLinesRange());
                    excelRowWriter.WriteCell(3, comment.Type);
                    excelRowWriter.WriteCell(4, comment.Metrics.WordsCount, comment.Evaluation.WordsCount());
                    excelRowWriter.WriteCell(5, comment.Metrics.HasNothing, comment.Metrics.HasNothing);
                    excelRowWriter.WriteCell(6, comment.Metrics.HasExclamationMark, comment.Metrics.HasExclamationMark);
                    excelRowWriter.WriteCell(7, comment.Metrics.HasQuestionMark, comment.Metrics.HasQuestionMark);
                    excelRowWriter.WriteCell(8, comment.Metrics.HasCode, comment.Metrics.HasCode);
                    excelRowWriter.WriteCell(9, comment.Metrics.CoherenceCoefficient, comment.Evaluation.CoherenceCoefficient());
                    excelRowWriter.WriteCell(10, comment.Metrics.LocationMethod);
                    excelRowWriter.WriteCell(11, comment.Metrics.MethodName);
                    excelRowWriter.WriteCell(12, comment.Metrics.LocationClass);
                    excelRowWriter.WriteCell(13, comment.Metrics.ClassName);
                    excelRowWriter.WriteCell(14, comment.Metrics.IsClassSmelly, comment.Metrics.IsClassSmelly);
                    excelRowWriter.WriteCell(15, comment.Evaluation.IsBad(), comment.Evaluation.IsBad());
                    excelRowWriter.WriteCell(16, comment.Content);

                    rowNumber++;
                }

                for (int i = 2; i <= 14; i++)
                {
                    worksheet.Column(i).AutoFit();
                }

                ExcelWorksheet summaryWorksheet = package.Workbook.Worksheets.Add("Summary");

                summaryWorksheet.Cells[1, 1].Value = "Percent of comments in smelly classes";
                summaryWorksheet.Cells[2, 1].Value =
                    $"{100.0 * commentStore.Comments.Count(c => c.Metrics.IsClassSmelly == true) / commentStore.Comments.Count}%";

                package.Save();
            }
        }
    }
}
