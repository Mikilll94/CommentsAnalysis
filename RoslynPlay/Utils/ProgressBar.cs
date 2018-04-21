using System;
using System.Collections.Generic;
using System.Text;

namespace RoslynPlay
{
    public class ProgressBar
    {
        private int _processedFiles = 0;
        private int _length;

        public ProgressBar(int length)
        {
            _length = length;
        }

        public void UpdateAndDisplay()
        {
            _processedFiles++;
            Console.Write($"\r{_processedFiles}/{_length}");
        }
    }
}
