using OfficeOpenXml;
using System.IO;
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
                new ClassesWorksheet(package, commentStore).Create();
                new RankingWorksheet(package).Create();
                new SummaryWorksheet(package, commentStore).Create();

                package.Save();
            }
        }
    }
}
