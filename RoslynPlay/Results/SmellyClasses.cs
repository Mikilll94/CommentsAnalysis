using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace RoslynPlay
{
    class SmellyClasses
    {
        public static HashSet<string> All { get; set; } = new HashSet<string>();
        public static HashSet<string> Abstraction { get; set; } = new HashSet<string>();
        public static HashSet<string> Encapsulation { get; set; } = new HashSet<string>();
        public static HashSet<string> Modularization { get; set; } = new HashSet<string>();
        public static HashSet<string> Hierarchy { get; set; } = new HashSet<string>();

        public SmellyClasses(string projectName, string sheetPrefix)
        {
            using (var stream = File.Open(projectName, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    Dictionary<string, HashSet<string>> sheets = new Dictionary<string, HashSet<string>>()
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
