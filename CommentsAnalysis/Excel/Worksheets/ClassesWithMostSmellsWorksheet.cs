using System.Linq;
using OfficeOpenXml;

namespace RoslynPlay
{
    public class ClassesWithMostSmellsWorksheet : Worksheet
    {
        private ClassStore _classStore;
        private CommentStore _commentStore;

        public ClassesWithMostSmellsWorksheet(ExcelPackage package, CommentStore commentStore, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Class";
            worksheet.Cells[1, 3].Value = "Smells count";
            worksheet.Cells[1, 4].Value = "Comments count";

            worksheet.Cells[1, 5].Value = "Non-documentation comments";
            worksheet.Cells[1, 6].Value = "Documentation comments";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classesWithMostSmells = _classStore.Classes.OrderByDescending(c => c.SmellsCount).ToArray();

            int rowNo = 2;

            foreach (var @class in classesWithMostSmells)
            {
                worksheet.Cells[rowNo, 1].Value = @class.FileName;
                worksheet.Cells[rowNo, 2].Value = @class.Name;
                worksheet.Cells[rowNo, 3].Value = @class.SmellsCount;
                worksheet.Cells[rowNo, 4].Value =
                    _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName);

                worksheet.Cells[rowNo, 5].Value = int.Parse(_commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName
                        && c.Type == CommentType.SingleLine).ToString())
                        + int.Parse(_commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName
                        && c.Type == CommentType.MultiLine).ToString());
                worksheet.Cells[rowNo, 6].Value =
                    _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName
                        && c.Type == CommentType.Doc);

                rowNo++;
            }

            worksheet.View.FreezePanes(2, 1);
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 2; i <= 6; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }
    }
}
