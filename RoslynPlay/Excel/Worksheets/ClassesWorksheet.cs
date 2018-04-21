using System.Linq;
using OfficeOpenXml;

namespace RoslynPlay
{
    class ClassesWorksheet : Worksheet
    {
        public ClassesWorksheet(ExcelPackage package, CommentStore commentStore) : base(package, commentStore)
        {
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
            Class[] classes = ClassStore.Classes.ToArray();

            for (int i = 2; i < classes.Length; i++)
            {
                worksheet.Cells[i, 1].Value = classes[i].FileName;
                worksheet.Cells[i, 2].Value = classes[i].Name;
                worksheet.Cells[i, 3].Value = classes[i].SmellsCount;
                worksheet.Cells[i, 4].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName);
            }
        }
    }
}
