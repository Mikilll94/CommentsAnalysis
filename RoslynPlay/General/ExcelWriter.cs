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
                    ExcelUtils.WriteCell(rowNumber, 1, comment.FileName, worksheet);
                    ExcelUtils.WriteCell(rowNumber, 2, comment.GetLinesRange(), worksheet);
                    ExcelUtils.WriteCell(rowNumber, 3, comment.Type, worksheet);
                    ExcelUtils.WriteCell(rowNumber, 4, comment.Metrics.WordsCount, worksheet, comment.Evaluation.WordsCount());
                    ExcelUtils.WriteCell(rowNumber, 5, comment.Metrics.HasNothing, worksheet, comment.Metrics.HasNothing);
                    ExcelUtils.WriteCell(rowNumber, 6, comment.Metrics.HasExclamationMark, worksheet, comment.Metrics.HasExclamationMark);
                    ExcelUtils.WriteCell(rowNumber, 7, comment.Metrics.HasQuestionMark, worksheet, comment.Metrics.HasQuestionMark);
                    ExcelUtils.WriteCell(rowNumber, 8, comment.Metrics.HasCode, worksheet, comment.Metrics.HasCode);
                    ExcelUtils.WriteCell(rowNumber, 9, comment.Metrics.CoherenceCoefficient, worksheet, comment.Evaluation.CoherenceCoefficient());
                    ExcelUtils.WriteCell(rowNumber, 10, comment.Metrics.LocationMethod, worksheet);
                    ExcelUtils.WriteCell(rowNumber, 11, comment.Metrics.MethodName, worksheet);
                    ExcelUtils.WriteCell(rowNumber, 12, comment.Metrics.LocationClass, worksheet);
                    ExcelUtils.WriteCell(rowNumber, 13, comment.Metrics.ClassName, worksheet);
                    ExcelUtils.WriteCell(rowNumber, 14, comment.Metrics.IsClassSmelly, worksheet, comment.Metrics.IsClassSmelly);
                    ExcelUtils.WriteCell(rowNumber, 15, comment.Evaluation.IsBad(), worksheet, comment.Evaluation.IsBad());
                    ExcelUtils.WriteCell(rowNumber, 16, comment.Content, worksheet);

                    rowNumber++;
                }

                for (int i = 2; i <= 14; i++)
                {
                    worksheet.Column(i).AutoFit();
                }
                package.Save();
            }
        }
    }
}
