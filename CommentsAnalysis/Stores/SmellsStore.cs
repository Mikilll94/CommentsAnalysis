using CommentsAnalysis.Models;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace CommentsAnalysis
{
    public class SmellsStore
    {
        public static List<Smell> Smells { get; set; } = new List<Smell>();

        public static void Initialize(string projectName, string sheetPrefix)
        {
            using (var stream = File.Open(projectName, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    string[] sheets = new string[] { "AbsSMells", "EncSMells", "ModSMells", "HieSMells" };

                    foreach (var sheet in sheets)
                    {
                        string sheetName = $"{sheetPrefix}_{sheet}";
                        DataRowCollection rows = result.Tables[sheetName].Rows;
                        for (int i = 0; i < rows.Count; i++)
                        {
                            string className = rows[i][2].ToString();
                            string @namespace = rows[i][1].ToString();

                            SmellType smellType;

                            switch (sheet)
                            {
                                case "AbsSMells":
                                    smellType = SmellType.Abstraction;
                                    break;
                                case "EncSMells":
                                    smellType = SmellType.Encapsulation;
                                    break;
                                case "ModSMells":
                                    smellType = SmellType.Modularization;
                                    break;
                                case "HieSMells":
                                    smellType = SmellType.Hierarchy;
                                    break;
                                default:
                                    throw new Exception("Unknown sheet name");
                            }

                            Smells.Add(new Smell() { Type = smellType, ClassName = className, ClassNamespace = @namespace });
                        }
                    }
                }
            }
        }
    }
}
