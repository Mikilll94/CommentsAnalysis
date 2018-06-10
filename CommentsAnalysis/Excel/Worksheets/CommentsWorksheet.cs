using OfficeOpenXml;

namespace CommentsAnalysis
{
    class CommentsWorksheet : Worksheet
    {
        private CommentStore _commentStore;

        public CommentsWorksheet(ExcelPackage package, CommentStore commentStore) : base(package)
        {
            _commentStore = commentStore;
        }

        protected override void WriteHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "File";
            worksheet.Cells[1, 2].Value = "Line";
            worksheet.Cells[1, 3].Value = "Type";
            worksheet.Cells[1, 4].Value = "Number of words";
            worksheet.Cells[1, 5].Value = "Has \"nothing\"";
            worksheet.Cells[1, 6].Value = "Has !";
            worksheet.Cells[1, 7].Value = "Has ?";
            worksheet.Cells[1, 8].Value = "Has code";
            worksheet.Cells[1, 9].Value = "Coherence coefficient";
            worksheet.Cells[1, 10].Value = "Location method";
            worksheet.Cells[1, 11].Value = "Method name";
            worksheet.Cells[1, 12].Value = "Location class";
            worksheet.Cells[1, 13].Value = "Class name";
            worksheet.Cells[1, 14].Value = "Is class smelly?";
            worksheet.Cells[1, 15].Value = "Is bad?";
            worksheet.Cells[1, 16].Value = "Comment";
        }

        protected override void WriteData(ExcelWorksheet worksheet)
        {
            int rowNumber = 2;

            foreach (var comment in _commentStore.Comments)
            {
                ExcelRowWriter excelRowWriter = new ExcelRowWriter(worksheet, rowNumber);

                excelRowWriter.WriteCell(1, comment.FileName);
                excelRowWriter.WriteCell(2, comment.GetLinesRange());
                excelRowWriter.WriteCell(3, comment.Type);
                excelRowWriter.WriteCell(4, comment.WordsCount, comment.IsBadWordsCount());
                excelRowWriter.WriteCell(5, comment.HasNothing, comment.HasNothing);
                excelRowWriter.WriteCell(6, comment.HasExclamationMark, comment.HasExclamationMark);
                excelRowWriter.WriteCell(7, comment.HasQuestionMark, comment.HasQuestionMark);
                excelRowWriter.WriteCell(8, comment.HasCode, comment.HasCode);
                excelRowWriter.WriteCell(9, comment.CoherenceCoefficient, comment.IsBadCoherenceCoefficient());
                excelRowWriter.WriteCell(10, comment.LocationRelativeToMethod);
                excelRowWriter.WriteCell(11, comment.MethodName);
                excelRowWriter.WriteCell(12, comment.LocationRelativeToClass);
                excelRowWriter.WriteCell(13, comment.Class?.Name);
                excelRowWriter.WriteCell(14, comment.Class?.IsSmelly(), comment.Class?.IsSmelly());
                excelRowWriter.WriteCell(15, comment.IsBad(), comment.IsBad());
                excelRowWriter.WriteCell(16, comment.Content);

                rowNumber++;
            }
        }

        protected override void FitColumns(ExcelWorksheet worksheet)
        {
            for (int i = 2; i <= 14; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }

        protected override void FreezePanes(ExcelWorksheet worksheet)
        {
            worksheet.View.FreezePanes(2, 1);
        }
    }
}
