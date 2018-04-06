using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RoslynPlay
{
    class SmellyClasses
    {
        public static List<string> Classes { get; set; } = new List<string>();

        public SmellyClasses(string projectName)
        {
            using (var stream = File.Open("c:/Users/wasni/Desktop/Designite_GitExtensions.xls", FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    for (int i = 0; i < result.Tables[$"{projectName}_AbsSMells"].Rows.Count; i++)
                    {
                        Classes.Add(result.Tables[$"{projectName}_AbsSMells"].Rows[i][2].ToString());
                    }
                    for (int i = 0; i < result.Tables[$"{projectName}_EncSMells"].Rows.Count; i++)
                    {
                        Classes.Add(result.Tables[$"{projectName}_EncSMells"].Rows[i][2].ToString());
                    }
                    for (int i = 0; i < result.Tables[$"{projectName}_ModSMells"].Rows.Count; i++)
                    {
                        Classes.Add(result.Tables[$"{projectName}_ModSMells"].Rows[i][2].ToString());
                    }
                    for (int i = 0; i < result.Tables[$"{projectName}_HieSMells"].Rows.Count; i++)
                    {
                        Classes.Add(result.Tables[$"{projectName}_HieSMells"].Rows[i][2].ToString());
                    }
                    Classes = Classes.Distinct().ToList();
                }
            }
        }
    }
}
