using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using CommentsAnalysis.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace RoslynPlay
{
    class ClassesWorksheet : Worksheet
    {
        private ClassStore _classStore;
        private CommentStore _commentStore;

        public ClassesWorksheet(ExcelPackage package, CommentStore commentStore, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Namespace";
            worksheet.Cells[1, 3].Value = "Class";
            worksheet.Cells[1, 4].Value = "Smells count";
            worksheet.Cells[1, 5].Value = "Comments count";

            worksheet.Cells[1, 6].Value = "Single line comments count";
            worksheet.Cells[1, 7].Value = "Multi line comments count";
            worksheet.Cells[1, 8].Value = "Non-documentation comments count";
            worksheet.Cells[1, 9].Value = "Documentation comments count";

            worksheet.Cells[1, 10].Value = "Method description";
            worksheet.Cells[1, 11].Value = "Method start";
            worksheet.Cells[1, 12].Value = "Method inner";
            worksheet.Cells[1, 13].Value = "Method end";

            worksheet.Cells[1, 14].Value = "Abstraction smells count";
            worksheet.Cells[1, 15].Value = "Encapsulation smells count";
            worksheet.Cells[1, 16].Value = "Modularization smells count";
            worksheet.Cells[1, 17].Value = "Hierarchy smells count";

            worksheet.Cells["T2"].Value = "Correlation between smells count and ...";
            worksheet.Cells["T2:V2"].Merge = true;
            worksheet.Cells["T3"].Value = "Comments";
            worksheet.Cells["U3"].Value = "Non-doc";
            worksheet.Cells["V3"].Value = "Doc";

            worksheet.Cells["T6"].Value = "Correlation between comments and ...";
            worksheet.Cells["T6:W6"].Merge = true;
            worksheet.Cells["T7"].Value = "Abstraction";
            worksheet.Cells["U7"].Value = "Encapsulation";
            worksheet.Cells["V7"].Value = "Modularization";
            worksheet.Cells["W7"].Value = "Hierarchy";

            SetColumnsColor(worksheet, Color.LightGoldenrodYellow, 6, 7, 8, 9);
            SetColumnsColor(worksheet, Color.LightGreen, 10, 11, 12, 13);
            SetColumnsColor(worksheet, Color.LightBlue, 14, 15, 16, 17);
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classes = _classStore.Classes.ToArray();

            List<double> smellsCountList = new List<double>();
            List<double> commentsCountList = new List<double>();
            List<double> nonDocCommentsCountList = new List<double>();
            List<double> docCommentsCountList = new List<double>();

            int rowNumber = 2;

            foreach (var @class in classes)
            {
                worksheet.Cells[rowNumber, 1].Value = @class.FileName;
                worksheet.Cells[rowNumber, 2].Value = @class.Namespace;
                worksheet.Cells[rowNumber, 3].Value = @class.Name;
                worksheet.Cells[rowNumber, 4].Value = @class.SmellsCount;

                Func<Comment, bool> classPredicate = c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName;

                worksheet.Cells[rowNumber, 5].Value = _commentStore.Comments.Where(classPredicate).Count();
                worksheet.Cells[rowNumber, 6].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.SingleLine);
                worksheet.Cells[rowNumber, 7].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.MultiLine);

                int nonDocCommentsCount = _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.SingleLine) + _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.MultiLine);
                worksheet.Cells[rowNumber, 8].Value = nonDocCommentsCount;

                worksheet.Cells[rowNumber, 9].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.Doc);

                worksheet.Cells[rowNumber, 10].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodDescription);
                worksheet.Cells[rowNumber, 11].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodStart);
                worksheet.Cells[rowNumber, 12].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodInner);
                worksheet.Cells[rowNumber, 13].Value = _commentStore.Comments.Where(classPredicate).Count(c => c.Metrics.LocationRelativeToMethod == LocationRelativeToMethod.MethodEnd);

                worksheet.Cells[rowNumber, 14].Value = @class.AbstractionSmellsCount;
                worksheet.Cells[rowNumber, 15].Value = @class.EncapsulationSmellsCount;
                worksheet.Cells[rowNumber, 16].Value = @class.ModularizationSmellsCount;
                worksheet.Cells[rowNumber, 17].Value = @class.HierarchySmellsCount;

                smellsCountList.Add(@class.SmellsCount);
                commentsCountList.Add(_commentStore.Comments.Where(classPredicate).Count());
                nonDocCommentsCountList.Add(nonDocCommentsCount);
                docCommentsCountList.Add(_commentStore.Comments.Where(classPredicate).Count(c => c.Type == CommentType.Doc));

                rowNumber++;
            }

            worksheet.Cells["T4"].Value = Math.Round(PearsonCorrelation.Compute(smellsCountList, commentsCountList), 3);
            worksheet.Cells["U4"].Value = Math.Round(PearsonCorrelation.Compute(smellsCountList, nonDocCommentsCountList), 3);
            worksheet.Cells["V4"].Value = Math.Round(PearsonCorrelation.Compute(smellsCountList, docCommentsCountList), 3);

            worksheet.Cells["T8"].Value = Math.Round(PearsonCorrelation.Compute(commentsCountList, classes.Select(c => (double)c.AbstractionSmellsCount).ToList()), 3);
            worksheet.Cells["U8"].Value = Math.Round(PearsonCorrelation.Compute(commentsCountList, classes.Select(c => (double)c.EncapsulationSmellsCount).ToList()), 3);
            worksheet.Cells["V8"].Value = Math.Round(PearsonCorrelation.Compute(commentsCountList, classes.Select(c => (double)c.ModularizationSmellsCount).ToList()), 3);
            worksheet.Cells["W8"].Value = Math.Round(PearsonCorrelation.Compute(commentsCountList, classes.Select(c => (double)c.HierarchySmellsCount).ToList()), 3);
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 4; i <= 17; i++)
            {
                worksheet.Column(i).AutoFit();
            }
            for (int i = 20; i <= 22; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }

        private void SetColumnsColor(ExcelWorksheet worksheet, Color color,  params int[] columnNumbers)
        {
            foreach (var column in columnNumbers)
            {
                worksheet.Column(column).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Column(column).Style.Fill.BackgroundColor.SetColor(color);
            }
        }

        protected override void FreezePanes(ExcelWorksheet worksheet)
        {
            worksheet.View.FreezePanes(2, 1);
        }
    }
}
