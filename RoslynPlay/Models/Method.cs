using System;
using System.Collections.Generic;
using System.Text;

namespace RoslynPlay
{
    class Method
    {
        public string Name { get; set; }
        public string FileName { get; set; }
        public int LineNumber { get; set; }
        public int LineEnd { get; set; } = -1;
    }
}
