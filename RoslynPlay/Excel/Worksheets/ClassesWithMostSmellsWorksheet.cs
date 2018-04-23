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
                rowNo++;
            }
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= 4; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }
    }
}
