using OfficeOpenXml;
using System;
using System.IO;
using System.Text;

namespace CommentsAnalysis
{
    public class ExcelWriter
    {
        private FileInfo _file;

        public ExcelWriter(string file)
        {
            _file = new FileInfo(file);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public void Write(CommentStore commentStore, ClassStore classStore)
        {
            using (ExcelPackage package = new ExcelPackage(_file))
            {
                new CommentsWorksheet(package, commentStore).Create("Comments");
                Console.WriteLine("Generated comments worksheet");
                new ClassesWorksheet(package, classStore).Create("Classes");
                Console.WriteLine("Generated classes worksheet");
                new ClassesWithMostSmellsWorksheet(package, commentStore, classStore).Create("ClassesWithMostSmells");
                Console.WriteLine("Generated classes with most smells worksheet");
                new ClassesWithMostComments(package, commentStore, classStore).Create("ClassesWithMostComments");
                Console.WriteLine("Generated classes with most comments worksheet");

                package.Save();
            }
        }
    }
}
