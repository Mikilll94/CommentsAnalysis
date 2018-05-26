using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace RoslynPlay
{
    public class SmellyClasses
    {
        public static List<Class> All { get; set; } = new List<Class>();
        public static List<Class> Abstraction { get; set; } = new List<Class>();
        public static List<Class> Encapsulation { get; set; } = new List<Class>();
        public static List<Class> Modularization { get; set; } = new List<Class>();
        public static List<Class> Hierarchy { get; set; } = new List<Class>();

        public SmellyClasses(string projectName, string sheetPrefix)
        {
            using (var stream = File.Open(projectName, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    Dictionary<string, List<Class>> sheets = new Dictionary<string, List<Class>>()
                    {
                        { "AbsSMells", Abstraction },
                        { "EncSMells", Encapsulation },
                        { "ModSMells", Modularization },
                        { "HieSMells", Hierarchy }
                    };

                    foreach (var sheet in sheets.Keys)
                    {
                        string sheetName = $"{sheetPrefix}_{sheet}";
                        DataRowCollection rows = result.Tables[sheetName].Rows;
                        for (int i = 0; i < rows.Count; i++)
                        {
                            string className = rows[i][2].ToString();
                            string @namespace = rows[i][1].ToString();
                            var @class = new Class() { Name = className, Namespace = @namespace };
                            All.Add(@class);
                            sheets[sheet].Add(@class);
                        }
                    }
                }
            }
        }
    }
}
