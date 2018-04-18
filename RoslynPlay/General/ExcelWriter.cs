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

                #region "Classes" sheet

                ExcelWorksheet classesWorksheet = package.Workbook.Worksheets.Add("Classes");

                classesWorksheet.Cells[1, 1].Value = "File name";
                classesWorksheet.Cells[1, 2].Value = "Class";
                classesWorksheet.Cells[1, 3].Value = "Smells count";
                classesWorksheet.Cells[1, 4].Value = "Comments count";

                Class[] classes = ClassStore.Classes.ToArray();

                for (int i = 2; i < classes.Length; i++)
                {
                    classesWorksheet.Cells[i, 1].Value = classes[i].FileName;
                    classesWorksheet.Cells[i, 2].Value = classes[i].Name;
                    classesWorksheet.Cells[i, 3].Value = classes[i].SmellsCount;
                    classesWorksheet.Cells[i, 4].Value = 
                        commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName);
                }

                #endregion

                ExcelWorksheet summaryWorksheet = package.Workbook.Worksheets.Add("Summary");

                summaryWorksheet.Cells[1, 1].Value = "Number of comments";
                summaryWorksheet.Cells[1, 1, 1, 6].Merge = true;

                summaryWorksheet.Cells[2, 1].Value = "Total";
                summaryWorksheet.Cells[2, 1, 3, 1].Merge = true;
                summaryWorksheet.Cells[4, 1].Value = commentStore.Comments.Count;

                summaryWorksheet.Cells[2, 3].Value = "In smelly classes";
                summaryWorksheet.Cells[2, 3, 2, 8].Merge = true;
                summaryWorksheet.Cells[3, 2].Value = "Bad";
                summaryWorksheet.Cells[4, 2].Value = commentStore.Comments.Count(c => c.Evaluation.IsBad() == true);
                summaryWorksheet.Cells[3, 3].Value = "Bad";
                summaryWorksheet.Cells[4, 3].Value = commentStore.Comments.Count(c => c.Evaluation.IsBad() == true && c.Metrics.IsClassSmelly == true);
                summaryWorksheet.Cells[3, 4].Value = "Abstraction";
                summaryWorksheet.Cells[4, 4].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyAbstraction == true);
                summaryWorksheet.Cells[3, 5].Value = "Encapsulation";
                summaryWorksheet.Cells[4, 5].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyEncapsulation == true);
                summaryWorksheet.Cells[3, 6].Value = "Modularization";
                summaryWorksheet.Cells[4, 6].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyModularization == true);
                summaryWorksheet.Cells[3, 7].Value = "Hierarchy";
                summaryWorksheet.Cells[4, 7].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyHierarchy == true);
                summaryWorksheet.Cells[3, 8].Value = "Total";
                summaryWorksheet.Cells[4, 8].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmelly == true);

                for (int i = 1; i <= 6; i++)
                {
                    summaryWorksheet.Column(i).AutoFit();
                    summaryWorksheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                package.Save();
            }
        }
    }
}
