using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace RoslynPlay
{
    public static class ExcelUtils
    {
        public static void WriteCell(int row, int column, object value, ExcelWorksheet worksheet, bool? formatCondition = null)
        {
            ExcelRange _excelRange = worksheet.Cells[row, column];
            _excelRange.Value = value;
            if (formatCondition != null)
            {
                _excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                _excelRange.Style.Fill.BackgroundColor.SetColor((bool)formatCondition ? Color.IndianRed : Color.LightGreen);
            }
        }
    }
}
