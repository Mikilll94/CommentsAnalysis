using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace RoslynPlay
{
    public class ExcelRowWriter
    {
        private ExcelWorksheet _worksheet;
        private int _rowNo;

        public ExcelRowWriter(ExcelWorksheet worksheet, int rowNo)
        {
            _worksheet = worksheet;
            _rowNo = rowNo;
        }

        public void WriteCell(int column, object value, bool? formatCondition = null)
        {
            ExcelRange _excelRange = _worksheet.Cells[_rowNo, column];
            _excelRange.Value = value;
            if (formatCondition != null)
            {
                _excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                _excelRange.Style.Fill.BackgroundColor.SetColor((bool)formatCondition ? Color.IndianRed : Color.LightGreen);
            }
        }
    }
}
