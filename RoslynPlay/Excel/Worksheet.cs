using OfficeOpenXml;

namespace RoslynPlay
{
    public class Worksheet
    {
        protected ExcelPackage _package;
        protected CommentStore _commentStore;

        public Worksheet(ExcelPackage package, CommentStore commentStore = null)
        {
            _package = package;
            _commentStore = commentStore;
        }

        public void Create(string worksheetName)
        {
            ExcelWorksheet worksheet = _package.Workbook.Worksheets.Add(worksheetName);
            worksheet.View.FreezePanes(2, 1);
            WriteHeaders(worksheet);
            WriteData(worksheet);
            FitColumns(worksheet);
        }

        protected virtual void WriteHeaders(ExcelWorksheet worksheet) {}
        protected virtual void WriteData(ExcelWorksheet worksheet) {}
        protected virtual void FitColumns(ExcelWorksheet worksheet) {}
    }
}
