using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CommentsAnalysis
{
    class ClassesWithMostComments : Worksheet
    {
        private ClassStore _classStore;

        public ClassesWithMostComments(ExcelPackage package, ClassStore classStore) : base(package)
        {
            _classStore = classStore;
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
            Class[] classes = _classStore.Classes.OrderByDescending(c => c.Comments.Count()).ToArray();

            int rowNo = 2;

            foreach (var @class in classes)
            {
                worksheet.Cells[rowNo, 1].Value = @class.FileName;
                worksheet.Cells[rowNo, 2].Value = @class.Name;
                worksheet.Cells[rowNo, 3].Value = @class.Comments.Count();
                worksheet.Cells[rowNo, 4].Value = @class.SmellsCount;
                worksheet.Cells[rowNo, 5].Value = @class.AbstractionSmellsCount;
                worksheet.Cells[rowNo, 6].Value = @class.EncapsulationSmellsCount;
                worksheet.Cells[rowNo, 7].Value = @class.ModularizationSmellsCount;
                worksheet.Cells[rowNo, 8].Value = @class.HierarchySmellsCount;

                rowNo++;
            }

            IEnumerable<Class> classesWithMostComments = _classStore.Classes.Where(c => c.Comments.Count() >= 10);
            IEnumerable<Class> restOfClasses = _classStore.Classes.Where(c => c.Comments.Count() < 10);

            worksheet.Cells["K5"].Value = Math.Round(classesWithMostComments.Average(c => c.SmellsCount), 3);
            worksheet.Cells["L5"].Value = Math.Round(classesWithMostComments.Average(c => c.AbstractionSmellsCount), 3);
            worksheet.Cells["M5"].Value = Math.Round(classesWithMostComments.Average(c => c.EncapsulationSmellsCount), 3);
            worksheet.Cells["N5"].Value = Math.Round(classesWithMostComments.Average(c => c.ModularizationSmellsCount), 3);
            worksheet.Cells["O5"].Value = Math.Round(classesWithMostComments.Average(c => c.HierarchySmellsCount), 3);

            worksheet.Cells["K6"].Value = Math.Round(restOfClasses.Average(c => c.SmellsCount), 3);
            worksheet.Cells["L6"].Value = Math.Round(restOfClasses.Average(c => c.AbstractionSmellsCount), 3);
            worksheet.Cells["M6"].Value = Math.Round(restOfClasses.Average(c => c.EncapsulationSmellsCount), 3);
            worksheet.Cells["N6"].Value = Math.Round(restOfClasses.Average(c => c.ModularizationSmellsCount), 3);
            worksheet.Cells["O6"].Value = Math.Round(restOfClasses.Average(c => c.HierarchySmellsCount), 3);
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
