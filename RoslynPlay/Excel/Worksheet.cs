using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoslynPlay
{
    public abstract class Worksheet
    {
        public abstract void Create();
        protected abstract void WriteHeaders(ExcelWorksheet worksheet);
        protected abstract void WriteData(ExcelWorksheet worksheet);
        protected abstract void FitColumns(ExcelWorksheet worksheet);
    }
}
