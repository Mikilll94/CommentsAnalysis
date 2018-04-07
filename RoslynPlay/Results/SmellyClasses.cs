using ExcelDataReader;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RoslynPlay
{
    class SmellyClasses
    {
        public static List<string> All { get; set; } = new List<string>();
        public static List<string> Abstraction { get; set; } = new List<string>();
        public static List<string> Encapsulation { get; set; } = new List<string>();
        public static List<string> Modularization { get; set; } = new List<string>();
        public static List<string> Hierarchy { get; set; } = new List<string>();

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
                        All.Add(result.Tables[$"{projectName}_AbsSMells"].Rows[i][2].ToString());
                        Abstraction.Add(result.Tables[$"{projectName}_AbsSMells"].Rows[i][2].ToString());
                    }
                    for (int i = 0; i < result.Tables[$"{projectName}_EncSMells"].Rows.Count; i++)
                    {
                        All.Add(result.Tables[$"{projectName}_EncSMells"].Rows[i][2].ToString());
                        Encapsulation.Add(result.Tables[$"{projectName}_EncSMells"].Rows[i][2].ToString());
                    }
                    for (int i = 0; i < result.Tables[$"{projectName}_ModSMells"].Rows.Count; i++)
                    {
                        All.Add(result.Tables[$"{projectName}_ModSMells"].Rows[i][2].ToString());
                        Modularization.Add(result.Tables[$"{projectName}_ModSMells"].Rows[i][2].ToString());
                    }
                    for (int i = 0; i < result.Tables[$"{projectName}_HieSMells"].Rows.Count; i++)
                    {
                        All.Add(result.Tables[$"{projectName}_HieSMells"].Rows[i][2].ToString());
                        Hierarchy.Add(result.Tables[$"{projectName}_HieSMells"].Rows[i][2].ToString());
                    }
                    All = All.Distinct().ToList();
                    Abstraction = Abstraction.Distinct().ToList();
                    Encapsulation = Encapsulation.Distinct().ToList();
                    Modularization = Modularization.Distinct().ToList();
                    Hierarchy = Hierarchy.Distinct().ToList();
                }
            }
        }
    }
}
