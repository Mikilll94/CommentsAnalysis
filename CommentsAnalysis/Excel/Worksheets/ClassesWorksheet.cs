using System.Drawing;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace RoslynPlay
{
    class ClassesWorksheet : Worksheet
    {
        private ClassStore _classStore;
        private CommentStore _commentStore;

        public ClassesWorksheet(ExcelPackage package, CommentStore commentStore, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Namespace";
            worksheet.Cells[1, 3].Value = "Class";
            worksheet.Cells[1, 4].Value = "Smells count";
            worksheet.Cells[1, 5].Value = "Comments count";

            worksheet.Cells[1, 6].Value = "Single line comments count";
            worksheet.Cells[1, 7].Value = "Multi line comments count";
            worksheet.Cells[1, 8].Value = "Non-documentation comments count";
            worksheet.Cells[1, 9].Value = "Documentation comments count";

            worksheet.Cells[1, 10].Value = "Method description";
            worksheet.Cells[1, 11].Value = "Method start";
            worksheet.Cells[1, 12].Value = "Method inner";
            worksheet.Cells[1, 13].Value = "Method end";

            worksheet.Cells[1, 14].Value = "Abstraction smells count";
            worksheet.Cells[1, 15].Value = "Encapsulation smells count";
            worksheet.Cells[1, 16].Value = "Modularization smells count";
            worksheet.Cells[1, 17].Value = "Hierarchy smells count";

            SetColumnsColor(worksheet, Color.LightGoldenrodYellow, 6, 7, 8, 9);
            SetColumnsColor(worksheet, Color.LightGreen, 10, 11, 12, 13);
            SetColumnsColor(worksheet, Color.LightBlue, 14, 15, 16, 17);
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classes = _classStore.Classes.ToArray();

            for (int i = 2; i < classes.Length; i++)
            {
                worksheet.Cells[i, 1].Value = classes[i].FileName;
                worksheet.Cells[i, 2].Value = classes[i].Namespace;
                worksheet.Cells[i, 3].Value = classes[i].Name;
                worksheet.Cells[i, 4].Value = classes[i].SmellsCount;

                worksheet.Cells[i, 5].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName);

                worksheet.Cells[i, 6].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Type == CommentType.SingleLine);
                worksheet.Cells[i, 7].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Type == CommentType.MultiLine);
                worksheet.Cells[i, 8].Value = int.Parse(worksheet.Cells[i, 5].Value.ToString()) + int.Parse(worksheet.Cells[i, 6].Value.ToString());
                worksheet.Cells[i, 9].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Type == CommentType.Doc);

                worksheet.Cells[i, 10].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodDescription);
                worksheet.Cells[i, 11].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodStart);
                worksheet.Cells[i, 12].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodInner);
                worksheet.Cells[i, 13].Value =
                    _commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName
                        && c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodEnd);

                worksheet.Cells[i, 14].Value = classes[i].AbstractionSmellsCount;
                worksheet.Cells[i, 15].Value = classes[i].EncapsulationSmellsCount;
                worksheet.Cells[i, 16].Value = classes[i].ModularizationSmellsCount;
                worksheet.Cells[i, 17].Value = classes[i].HierarchySmellsCount;
            }
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 3; i <= 17; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }

        private void SetColumnsColor(ExcelWorksheet worksheet, Color color,  params int[] columnNumbers)
        {
            foreach (var column in columnNumbers)
            {
                worksheet.Column(column).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Column(column).Style.Fill.BackgroundColor.SetColor(color);
            }
        }

        protected override void FreezePanes(ExcelWorksheet worksheet)
        {
            worksheet.View.FreezePanes(2, 1);
        }
    }
}
