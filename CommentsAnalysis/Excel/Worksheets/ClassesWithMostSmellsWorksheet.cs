using System;
using System.Linq;
using OfficeOpenXml;

namespace CommentsAnalysis
{
    public class ClassesWithMostSmellsWorksheet : Worksheet
    {
        private ClassStore _classStore;
        private CommentStore _commentStore;

        public ClassesWithMostSmellsWorksheet(ExcelPackage package, CommentStore commentStore, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Class";
            worksheet.Cells[1, 3].Value = "Smells count";
            worksheet.Cells[1, 4].Value = "Comments count";

            worksheet.Cells[1, 5].Value = "Non-documentation comments";
            worksheet.Cells[1, 6].Value = "Documentation comments";

            worksheet.Cells["I2"].Value = "Average number of comments";
            worksheet.Cells["I2:K2"].Merge = true;
            worksheet.Cells["I3"].Value = "All";
            worksheet.Cells["J3"].Value = "Non-doc";
            worksheet.Cells["K3"].Value = "Doc";

            worksheet.Cells["H4"].Value = ">= 3 smells";
            worksheet.Cells["H5"].Value = "< 3 smells";

            worksheet.Cells["I7"].Value = "P-value";
            worksheet.Cells["I7:K7"].Merge = true;
            worksheet.Cells["I8"].Value = "All";
            worksheet.Cells["J8"].Value = "Non-doc";
            worksheet.Cells["K8"].Value = "Doc";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classesWithMostSmells = _classStore.Classes.OrderByDescending(c => c.SmellsCount).ToArray();

            int rowNo = 2;

            int commentsCountInSmellyClasses = 0;
            int nonDocCommentsCountInSmellyClasses = 0;
            int docCommentsCountInSmellyClasses = 0;
            int smellyClassesCount = 0;

            int commentsCountInCleanClasses = 0;
            int nonDocCommentsCountInCleanClasses = 0;
            int docCommentsCountInCleanClasses = 0;
            int cleanClassesCount = 0;

            foreach (var @class in classesWithMostSmells)
            {
                worksheet.Cells[rowNo, 1].Value = @class.FileName;
                worksheet.Cells[rowNo, 2].Value = @class.Name;
                worksheet.Cells[rowNo, 3].Value = @class.SmellsCount;

                Func<Comment, bool> classPredicate = c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName;

                worksheet.Cells[rowNo, 4].Value = _commentStore.Comments.Where(classPredicate).Count();

                int nonDocCommentsCount = _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.SingleLine)
                        + _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.MultiLine);
                worksheet.Cells[rowNo, 5].Value = nonDocCommentsCount;

                worksheet.Cells[rowNo, 6].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.Doc);

                if (@class.SmellsCount >= 3)
                {
                    commentsCountInSmellyClasses += _commentStore.Comments.Where(classPredicate).Count();
                    nonDocCommentsCountInSmellyClasses += nonDocCommentsCount;
                    docCommentsCountInSmellyClasses += _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.Doc);
                    smellyClassesCount++;
                }
                else
                {
                    commentsCountInCleanClasses += _commentStore.Comments.Where(classPredicate).Count();
                    nonDocCommentsCountInCleanClasses += nonDocCommentsCount;
                    docCommentsCountInCleanClasses += _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.Doc);
                    cleanClassesCount++;
                }

                rowNo++;
            }

            worksheet.Cells["I4"].Value = Math.Round((double)commentsCountInSmellyClasses / smellyClassesCount,3);
            worksheet.Cells["J4"].Value = Math.Round((double)nonDocCommentsCountInSmellyClasses / smellyClassesCount, 3);
            worksheet.Cells["K4"].Value = Math.Round((double)docCommentsCountInSmellyClasses / smellyClassesCount, 3);

            worksheet.Cells["I5"].Value = Math.Round((double)commentsCountInCleanClasses / cleanClassesCount, 3);
            worksheet.Cells["J5"].Value = Math.Round((double)nonDocCommentsCountInCleanClasses / cleanClassesCount, 3);
            worksheet.Cells["K5"].Value = Math.Round((double)docCommentsCountInCleanClasses / cleanClassesCount, 3);

            //worksheet.Cells["I8"].Formula = "TTEST(B2:B117;B118:B806;1,3)";
            //worksheet.Cells["J8"].Value = Math.Round((double)nonDocCommentsCountInCleanClasses / cleanClassesCount, 3);
            //worksheet.Cells["K8"].Value = Math.Round((double)docCommentsCountInCleanClasses / cleanClassesCount, 3);
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 2; i <= 6; i++)
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
