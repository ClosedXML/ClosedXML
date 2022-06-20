using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class InsertingDeletingColumns : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Inserting and Deleting Columns");

            // Range starts with 2 columns
            var rng = ws.Range("B2:C3"); // Range starts on B2

            // Insert a column before the range
            ws.Column(1).InsertColumnsAfter(1); // Range starts on C2

            // Insert a column in between the range
            ws.Column(3).InsertColumnsAfter(1); // Range now has 3 columns

            // Insert a column (from a range) before the range
            ws.Range("A2:A3").InsertColumnsAfter(1); // Range starts on D2

            // Insert a column (from a range) in between the range
            ws.Range("D2:D3").InsertColumnsAfter(1); // Range now has 4 columns

            // Inserting columns from a range not covering all columns
            // does not affect our defined range
            ws.Range("A1:A2").InsertColumnsAfter(1);
            ws.Range("E3:E4").InsertColumnsAfter(1);

            // Delete a column before the range
            ws.Column(1).Delete(); // Range starts on C2

            // Delete a column (from a range) before the range
            ws.Range("A2:A3").Delete(XLShiftDeletedCells.ShiftCellsLeft); // Range starts on B2

            // Delete a column in between the range
            ws.Column(3).Delete(); // Range now has 3 columns

            // Delete a column (from a range) in between the range
            ws.Range("C2:C3").Delete(XLShiftDeletedCells.ShiftCellsLeft); // Range now has 2 columns

            // Deleting columns from a range not covering all columns
            // does not affect our defined range
            ws.Range("A1:A2").Delete(XLShiftDeletedCells.ShiftCellsLeft);
            ws.Range("D3:D4").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            rng.Style.Fill.BackgroundColor = XLColor.Orange;

            workbook.SaveAs(filePath);
        }
    }
}