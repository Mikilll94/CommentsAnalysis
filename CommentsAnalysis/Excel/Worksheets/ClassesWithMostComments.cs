using OfficeOpenXml;
using System.Linq;

namespace RoslynPlay
{
    class ClassesWithMostComments : Worksheet
    {
        private ClassStore _classStore;
        private CommentStore _commentStore;

        public ClassesWithMostComments(ExcelPackage package, CommentStore commentStore, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Class";
            worksheet.Cells[1, 3].Value = "Comments count";
            worksheet.Cells[1, 4].Value = "Smells count";
            worksheet.Cells[1, 5].Value = "Abstraction smells count";
            worksheet.Cells[1, 6].Value = "Encapsulation smells count";
            worksheet.Cells[1, 7].Value = "Modularization smells count";
            worksheet.Cells[1, 8].Value = "Hierarchy smells count";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classes = _classStore.Classes.OrderByDescending(@class => 
                _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName)).ToArray();

            int rowNo = 2;

            foreach (var @class in classes)
            {
                worksheet.Cells[rowNo, 1].Value = @class.FileName;
                worksheet.Cells[rowNo, 2].Value = @class.Name;
                worksheet.Cells[rowNo, 3].Value =
                    _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName);
                worksheet.Cells[rowNo, 4].Value = @class.SmellsCount;
                worksheet.Cells[rowNo, 5].Value = @class.AbstractionSmellsCount;
                worksheet.Cells[rowNo, 6].Value = @class.EncapsulationSmellsCount;
                worksheet.Cells[rowNo, 7].Value = @class.ModularizationSmellsCount;
                worksheet.Cells[rowNo, 8].Value = @class.HierarchySmellsCount;

                rowNo++;
            }

            worksheet.View.FreezePanes(2, 1);
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 2; i <= 8; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }
    }
}
