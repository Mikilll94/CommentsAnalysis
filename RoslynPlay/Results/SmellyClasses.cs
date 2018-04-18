using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

        public SmellyClasses(string projectName, string sheetPrefix)
        {
            using (var stream = File.Open(projectName, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    Dictionary<string, List<string>> sheets = new Dictionary<string, List<string>>()
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
                            All.Add(className);
                            sheets[sheet].Add(className);
                        }
                    }
                }
            }
        }
    }
}
