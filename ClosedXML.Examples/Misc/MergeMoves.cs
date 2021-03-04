using System.IO;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Misc
{
    public class MergeMoves : IXLExample
    {

        public void Create(string filePath)
        {
            string tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new MergeCells().Create(tempFile);
                var workbook = new XLWorkbook(tempFile);

                var ws = workbook.Worksheet(1);

                ws.Range("B1:F1").InsertRowsBelow(1);
                ws.Range("A3:A9").InsertColumnsAfter(1);
                ws.Row(1).Delete();
                ws.Column(1).Delete();

                ws.Range("E8:E9").InsertColumnsAfter(1);
                ws.Range("F2:F8").Merge();
                ws.Range("E3:E4").InsertColumnsAfter(1);
                ws.Range("F2:F8").Merge();
                ws.Range("E1:E2").InsertColumnsAfter(1);
                ws.Range("G2:G8").Merge();
                ws.Range("E1:E2").Delete(XLShiftDeletedCells.ShiftCellsLeft);

                ws.Range("D3:E3").InsertRowsBelow(1);
                ws.Range("A1:B1").InsertRowsBelow(1);
                ws.Range("B3:D3").Merge();
                ws.Range("A1:B1").Delete(XLShiftDeletedCells.ShiftCellsUp);

                ws.Range("B8:D8").Merge();
                ws.Range("D8:D9").Clear();

                workbook.SaveAs(filePath);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

    }
}
