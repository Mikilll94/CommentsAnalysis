using OfficeOpenXml;

namespace RoslynPlay
{
    public class Worksheet
    {
        protected ExcelPackage _package;

        public Worksheet(ExcelPackage package)
        {
            _package = package;
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
