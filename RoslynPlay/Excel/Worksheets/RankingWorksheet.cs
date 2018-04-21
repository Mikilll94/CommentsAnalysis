using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace RoslynPlay
{
    public class RankingWorksheet : Worksheet
    {
        private ExcelPackage _package;

        public RankingWorksheet(ExcelPackage package)
        {
            _package = package;
        }

        public override void Create()
        {
            ExcelWorksheet worksheet = _package.Workbook.Worksheets.Add("Ranking");
            WriteHeaders(worksheet);
            WriteData(worksheet);
            FitColumns(worksheet);
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Class";
            worksheet.Cells[1, 2].Value = "File name";
            worksheet.Cells[1, 3].Value = "No. of smells";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classesWithMostSmells = ClassStore.Classes.OrderByDescending(c => c.SmellsCount).Take(10).ToArray();

            int rankingRowNumber = 2;

            foreach (var @class in classesWithMostSmells)
            {
                worksheet.Cells[rankingRowNumber, 1].Value = @class.Name;
                worksheet.Cells[rankingRowNumber, 2].Value = @class.FileName;
                worksheet.Cells[rankingRowNumber, 3].Value = @class.SmellsCount;
                rankingRowNumber++;
            }
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
        }
    }
}
