using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace RoslynPlay
{
    class SummaryWorksheet : Worksheet
    {
        private CommentStore _commentStore;

        public SummaryWorksheet(ExcelPackage package, CommentStore commentStore) : base(package)
        {
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1, 1, 8].Merge = true;
            worksheet.Cells[2, 1, 3, 1].Merge = true;
            worksheet.Cells[2, 3, 2, 8].Merge = true;
            worksheet.Cells["B2:B3"].Merge = true;

            worksheet.Cells[1, 1].Value = "Number of comments";
            worksheet.Cells[2, 1].Value = "Total";
            worksheet.Cells[2, 3].Value = "In smelly classes";
            worksheet.Cells[2, 2].Value = "Bad";
            worksheet.Cells[3, 3].Value = "Bad";
            worksheet.Cells[3, 4].Value = "Abstraction";
            worksheet.Cells[3, 5].Value = "Encapsulation";
            worksheet.Cells[3, 6].Value = "Modularization";
            worksheet.Cells[3, 7].Value = "Hierarchy";
            worksheet.Cells[3, 8].Value = "Total";

            SetBorder(worksheet, "A1:H1");
            SetBorder(worksheet, "C2:H2");

            worksheet.Cells["C3:C4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["C3:C4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["H3:H4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["H3:H4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["A4:H4"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            worksheet.Cells[4, 1].Value = _commentStore.Comments.Count;
            worksheet.Cells[4, 2].Value = _commentStore.Comments.Count(c => c.Evaluation.IsBad() == true);
            worksheet.Cells[4, 3].Value = _commentStore.Comments.Count(c => c.Evaluation.IsBad() == true && c.Metrics.IsClassSmelly == true);
            worksheet.Cells[4, 4].Value = _commentStore.Comments.Count(c => c.Metrics.IsClassSmellyAbstraction == true);
            worksheet.Cells[4, 5].Value = _commentStore.Comments.Count(c => c.Metrics.IsClassSmellyEncapsulation == true);
            worksheet.Cells[4, 6].Value = _commentStore.Comments.Count(c => c.Metrics.IsClassSmellyModularization == true);
            worksheet.Cells[4, 7].Value = _commentStore.Comments.Count(c => c.Metrics.IsClassSmellyHierarchy == true);
            worksheet.Cells[4, 8].Value = _commentStore.Comments.Count(c => c.Metrics.IsClassSmelly == true);
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= 6; i++)
            {
                worksheet.Column(i).AutoFit();
                worksheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }

        private void SetBorder(ExcelWorksheet worksheet, string cells)
        {
            worksheet.Cells[cells].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[cells].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[cells].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[cells].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        protected override void FreezePanes(ExcelWorksheet worksheet)
        {
        }
    }
}
