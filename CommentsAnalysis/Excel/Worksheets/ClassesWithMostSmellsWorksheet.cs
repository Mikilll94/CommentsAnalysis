using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace CommentsAnalysis
{
    public class ClassesWithMostSmellsWorksheet : Worksheet
    {
        private ClassStore _classStore;

        public ClassesWithMostSmellsWorksheet(ExcelPackage package, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Class";
            worksheet.Cells[1, 3].Value = "Smells count";
            worksheet.Cells[1, 4].Value = "Comments count";

            worksheet.Cells[1, 5].Value = "Non-documentation comments";
            worksheet.Cells[1, 6].Value = "Documentation comments";

            worksheet.Cells[1, 7].Value = "Bad comments count";
            worksheet.Cells[1, 8].Value = "Bad non-documentation comments count";
            worksheet.Cells[1, 9].Value = "Bad documentation comments count";

            worksheet.Cells["M2"].Value = "Average number of comments";
            worksheet.Cells["M2:O2"].Merge = true;
            worksheet.Cells["M3"].Value = "All";
            worksheet.Cells["N3"].Value = "Non-doc";
            worksheet.Cells["O3"].Value = "Doc";

            worksheet.Cells["L4"].Value = ">= 3 smells";
            worksheet.Cells["L5"].Value = "< 3 smells";

            worksheet.Cells["M7"].Value = "Number of bad comments";
            worksheet.Cells["M7:O7"].Merge = true;
            worksheet.Cells["M8"].Value = "All";
            worksheet.Cells["N8"].Value = "Non-doc";
            worksheet.Cells["O8"].Value = "Doc";

            worksheet.Cells["L9"].Value = "with smells";
            worksheet.Cells["L10"].Value = "no smells";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classesWithMostSmells = _classStore.Classes.OrderByDescending(c => c.SmellsCount).ToArray();

            int rowNo = 2;

            foreach (var @class in classesWithMostSmells)
            {
                worksheet.Cells[rowNo, 1].Value = @class.FileName;
                worksheet.Cells[rowNo, 2].Value = @class.Name;
                worksheet.Cells[rowNo, 3].Value = @class.SmellsCount;
                worksheet.Cells[rowNo, 4].Value = @class.Comments.Count();

                worksheet.Cells[rowNo, 5].Value = @class.Comments.Count(c => c.Type == CommentType.SingleLine || c.Type == CommentType.MultiLine);
                worksheet.Cells[rowNo, 6].Value = @class.Comments.Count(c => c.Type == CommentType.Doc);

                worksheet.Cells[rowNo, 7].Value = @class.Comments.Count(c => c.IsBad() == true);
                worksheet.Cells[rowNo, 8].Value = @class.Comments.Count(c => c.IsBad() == true && (c.Type == CommentType.SingleLine || c.Type == CommentType.MultiLine));
                worksheet.Cells[rowNo, 9].Value = @class.Comments.Count(c => c.IsBad() == true && c.Type == CommentType.Doc);

                rowNo++;
            }

            IEnumerable<Class> smellyClasses = _classStore.Classes.Where(c => c.SmellsCount >= 3);
            IEnumerable<Class> cleanClasses = _classStore.Classes.Where(c => c.SmellsCount < 3);

            IEnumerable<Class> classesWithSmells = _classStore.Classes.Where(c => c.SmellsCount > 0);
            IEnumerable<Class> classesWithoutSmells = _classStore.Classes.Where(c => c.SmellsCount == 0);

            worksheet.Cells["M4"].Value = Math.Round(smellyClasses.Average(c => c.Comments.Count()), 3);
            worksheet.Cells["N4"].Value = Math.Round(smellyClasses.Average(c => c.Comments.Count(com => com.Type == CommentType.SingleLine || com.Type == CommentType.MultiLine)), 3);
            worksheet.Cells["O4"].Value = Math.Round(smellyClasses.Average(c => c.Comments.Count(com => com.Type == CommentType.Doc)), 3);

            worksheet.Cells["M5"].Value = Math.Round(cleanClasses.Average(c => c.Comments.Count()), 3);
            worksheet.Cells["N5"].Value = Math.Round(cleanClasses.Average(c => c.Comments.Count(com => com.Type == CommentType.SingleLine || com.Type == CommentType.MultiLine)), 3);
            worksheet.Cells["O5"].Value = Math.Round(cleanClasses.Average(c => c.Comments.Count(com => com.Type == CommentType.Doc)), 3);

            worksheet.Cells["M9"].Value = classesWithSmells.Sum(c => c.Comments.Count(com => com.IsBad()));
            worksheet.Cells["N9"].Value = classesWithSmells.Sum(c => c.Comments.Count(com => com.IsBad() && (com.Type == CommentType.SingleLine || com.Type == CommentType.MultiLine)));
            worksheet.Cells["O9"].Value = classesWithSmells.Sum(c => c.Comments.Count(com => com.IsBad() && com.Type == CommentType.Doc));

            worksheet.Cells["M10"].Value = classesWithoutSmells.Sum(c => c.Comments.Count(com => com.IsBad()));
            worksheet.Cells["N10"].Value = classesWithoutSmells.Sum(c => c.Comments.Count(com => com.IsBad() && (com.Type == CommentType.SingleLine || com.Type == CommentType.MultiLine)));
            worksheet.Cells["O10"].Value = classesWithoutSmells.Sum(c => c.Comments.Count(com => com.IsBad() && com.Type == CommentType.Doc));
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 2; i <= 9; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }

        protected override void FreezePanes(ExcelWorksheet worksheet)
        {
            worksheet.View.FreezePanes(2, 1);
        }
    }
}
