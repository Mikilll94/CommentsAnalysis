using OfficeOpenXml;
using System;
using System.Collections.Generic;
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
            worksheet.Cells[1, 1].Value = "File name";
            worksheet.Cells[1, 2].Value = "Class";
            worksheet.Cells[1, 3].Value = "Comments count";
            worksheet.Cells[1, 4].Value = "Smells count";
            worksheet.Cells[1, 5].Value = "Abstraction smells count";
            worksheet.Cells[1, 6].Value = "Encapsulation smells count";
            worksheet.Cells[1, 7].Value = "Modularization smells count";
            worksheet.Cells[1, 8].Value = "Hierarchy smells count";

            worksheet.Cells["K3"].Value = "Average number of smells";
            worksheet.Cells["K3:O3"].Merge = true;

            worksheet.Cells["K4"].Value = "All";
            worksheet.Cells["L4"].Value = "Abstraction";
            worksheet.Cells["M4"].Value = "Encapsulation";
            worksheet.Cells["N4"].Value = "Modularization";
            worksheet.Cells["O4"].Value = "Hierarchy";

            worksheet.Cells["J5"].Value = ">= 10 comments";
            worksheet.Cells["J6"].Value = "< 10 comments";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            Class[] classes = _classStore.Classes.OrderByDescending(@class => 
                _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName)).ToArray();

            List<Class> classesWithMostComments = new List<Class>();
            List<Class> restOfClasses = new List<Class>();

            int rowNo = 2;

            foreach (var @class in classes)
            {
                int commentsCount = _commentStore.Comments.Count(c => @class.Name == c.Metrics.ClassName && @class.FileName == c.FileName);

                worksheet.Cells[rowNo, 1].Value = @class.FileName;
                worksheet.Cells[rowNo, 2].Value = @class.Name;
                worksheet.Cells[rowNo, 3].Value = commentsCount;
                worksheet.Cells[rowNo, 4].Value = @class.SmellsCount;
                worksheet.Cells[rowNo, 5].Value = @class.AbstractionSmellsCount;
                worksheet.Cells[rowNo, 6].Value = @class.EncapsulationSmellsCount;
                worksheet.Cells[rowNo, 7].Value = @class.ModularizationSmellsCount;
                worksheet.Cells[rowNo, 8].Value = @class.HierarchySmellsCount;

                if (commentsCount >= 10)
                    classesWithMostComments.Add(@class);
                else
                    restOfClasses.Add(@class);

                rowNo++;
            }

            worksheet.Cells["K5"].Value = Math.Round((double)classesWithMostComments.Select(c => c.SmellsCount).Sum() / classesWithMostComments.Count, 3);
            worksheet.Cells["L5"].Value = Math.Round((double)classesWithMostComments.Select(c => c.AbstractionSmellsCount).Sum() / classesWithMostComments.Count, 3);
            worksheet.Cells["M5"].Value = Math.Round((double)classesWithMostComments.Select(c => c.EncapsulationSmellsCount).Sum() / classesWithMostComments.Count, 3);
            worksheet.Cells["N5"].Value = Math.Round((double)classesWithMostComments.Select(c => c.ModularizationSmellsCount).Sum() / classesWithMostComments.Count, 3);
            worksheet.Cells["O5"].Value = Math.Round((double)classesWithMostComments.Select(c => c.HierarchySmellsCount).Sum() / classesWithMostComments.Count, 3);

            worksheet.Cells["K6"].Value = Math.Round((double)restOfClasses.Select(c => c.SmellsCount).Sum() / restOfClasses.Count, 3);
            worksheet.Cells["L6"].Value = Math.Round((double)restOfClasses.Select(c => c.AbstractionSmellsCount).Sum() / restOfClasses.Count, 3);
            worksheet.Cells["M6"].Value = Math.Round((double)restOfClasses.Select(c => c.EncapsulationSmellsCount).Sum() / restOfClasses.Count, 3);
            worksheet.Cells["N6"].Value = Math.Round((double)restOfClasses.Select(c => c.ModularizationSmellsCount).Sum() / restOfClasses.Count, 3);
            worksheet.Cells["O6"].Value = Math.Round((double)restOfClasses.Select(c => c.HierarchySmellsCount).Sum() / restOfClasses.Count, 3);
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 2; i <= 8; i++)
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
