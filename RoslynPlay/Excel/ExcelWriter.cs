using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Linq;
using System.Text;

namespace RoslynPlay
{
    class ExcelWriter
    {
        private FileInfo _file;

        public ExcelWriter(string file)
        {
            _file = new FileInfo(file);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public void Write(CommentStore commentStore)
        {
            using (ExcelPackage package = new ExcelPackage(_file))
            {
                new CommentsWorksheet(package, commentStore).Create();

                #region "Classes" sheet

                ExcelWorksheet classesWorksheet = package.Workbook.Worksheets.Add("Classes");

                classesWorksheet.Cells[1, 1].Value = "File name";
                classesWorksheet.Cells[1, 2].Value = "Class";
                classesWorksheet.Cells[1, 3].Value = "Smells count";
                classesWorksheet.Cells[1, 4].Value = "Comments count";

                Class[] classes = ClassStore.Classes.ToArray();

                for (int i = 2; i < classes.Length; i++)
                {
                    classesWorksheet.Cells[i, 1].Value = classes[i].FileName;
                    classesWorksheet.Cells[i, 2].Value = classes[i].Name;
                    classesWorksheet.Cells[i, 3].Value = classes[i].SmellsCount;
                    classesWorksheet.Cells[i, 4].Value = 
                        commentStore.Comments.Count(c => classes[i].Name == c.Metrics.ClassName && classes[i].FileName == c.FileName);
                }

                #endregion

                ExcelWorksheet rankingWorksheet = package.Workbook.Worksheets.Add("Ranking");

                Class[] classesWithMostSmells = classes.Distinct().OrderByDescending(c => c.SmellsCount).Take(10).ToArray();
                rankingWorksheet.Cells[1, 1].Value = "Class";
                rankingWorksheet.Cells[1, 2].Value = "File name";
                rankingWorksheet.Cells[1, 3].Value = "No. of smells";

                int rankingRowNumber = 2;

                foreach (var classWithSmells in classesWithMostSmells)
                {
                    rankingWorksheet.Cells[rankingRowNumber, 1].Value = classWithSmells.Name;
                    rankingWorksheet.Cells[rankingRowNumber, 2].Value = classWithSmells.FileName;
                    rankingWorksheet.Cells[rankingRowNumber, 3].Value = classWithSmells.SmellsCount;
                    rankingRowNumber++;
                }

                ExcelWorksheet summaryWorksheet = package.Workbook.Worksheets.Add("Summary");

                summaryWorksheet.Cells[1, 1].Value = "Number of comments";
                summaryWorksheet.Cells[1, 1, 1, 6].Merge = true;

                summaryWorksheet.Cells[2, 1].Value = "Total";
                summaryWorksheet.Cells[2, 1, 3, 1].Merge = true;
                summaryWorksheet.Cells[4, 1].Value = commentStore.Comments.Count;

                summaryWorksheet.Cells[2, 3].Value = "In smelly classes";
                summaryWorksheet.Cells[2, 3, 2, 8].Merge = true;
                summaryWorksheet.Cells[3, 2].Value = "Bad";
                summaryWorksheet.Cells[4, 2].Value = commentStore.Comments.Count(c => c.Evaluation.IsBad() == true);
                summaryWorksheet.Cells[3, 3].Value = "Bad";
                summaryWorksheet.Cells[4, 3].Value = commentStore.Comments.Count(c => c.Evaluation.IsBad() == true && c.Metrics.IsClassSmelly == true);
                summaryWorksheet.Cells[3, 4].Value = "Abstraction";
                summaryWorksheet.Cells[4, 4].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyAbstraction == true);
                summaryWorksheet.Cells[3, 5].Value = "Encapsulation";
                summaryWorksheet.Cells[4, 5].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyEncapsulation == true);
                summaryWorksheet.Cells[3, 6].Value = "Modularization";
                summaryWorksheet.Cells[4, 6].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyModularization == true);
                summaryWorksheet.Cells[3, 7].Value = "Hierarchy";
                summaryWorksheet.Cells[4, 7].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmellyHierarchy == true);
                summaryWorksheet.Cells[3, 8].Value = "Total";
                summaryWorksheet.Cells[4, 8].Value = commentStore.Comments.Count(c => c.Metrics.IsClassSmelly == true);

                for (int i = 1; i <= 6; i++)
                {
                    summaryWorksheet.Column(i).AutoFit();
                    summaryWorksheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                package.Save();
            }
        }
    }
}
