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
            worksheet.Cells[1, 1].Value = "Class";
            worksheet.Cells[1, 2].Value = "File name";
            worksheet.Cells[1, 3].Value = "Comments count";
            worksheet.Cells[1, 4].Value = "Smells count";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classes = _classStore.Classes.OrderByDescending(@class => 
                _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName)).ToArray();

            int rowNo = 2;

            foreach (var @class in classes)
            {
                worksheet.Cells[rowNo, 1].Value = @class.Name;
                worksheet.Cells[rowNo, 2].Value = @class.FileName;
                worksheet.Cells[rowNo, 3].Value =
                    _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName);
                worksheet.Cells[rowNo, 4].Value = @class.SmellsCount;

                rowNo++;
            }
        }
    }
}
