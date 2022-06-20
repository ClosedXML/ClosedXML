using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class InsertingDeletingRows : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Inserting and Deleting Rows");

            // Range starts with 2 rows
            var rng = ws.Range("B2:C3");

            // Insert a row above the range
            ws.Row(1).InsertRowsBelow(1); // Range starts on B3

            // Insert a row in between the range
            ws.Row(3).InsertRowsBelow(1); // Range now has 3 rows

            // Insert a row (from a range) above the range
            ws.Range("B1:C1").InsertRowsBelow(1); // Range starts on B4

            // Insert a row (from a range) in between the range
            ws.Range("B4:C4").InsertRowsBelow(1); // Range now has 4 rows

            // Inserting rows from a range not covering all columns
            // does not affect our defined range
            ws.Range("A1:B1").InsertRowsBelow(1);
            ws.Range("C4:D4").InsertRowsBelow(1);

            // Delete a row above the range
            ws.Row(1).Delete(); // Range starts on B3

            // Delete a row (from a range) above the range
            ws.Range("B1:C1").Delete(XLShiftDeletedCells.ShiftCellsUp); // Range starts on B2

            // Delete a row in between the range
            ws.Row(3).Delete(); // Range now has 3 rows

            // Delete a row (from a range) in between the range
            ws.Range("B3:C3").Delete(XLShiftDeletedCells.ShiftCellsUp); // Range now has 2 rows

            // Deleting rows from a range not covering all columns
            // does not affect our defined range
            ws.Range("A1:B1").Delete(XLShiftDeletedCells.ShiftCellsUp);
            ws.Range("C4:D4").Delete(XLShiftDeletedCells.ShiftCellsUp);

            rng.Style.Fill.BackgroundColor = XLColor.Orange;

            workbook.SaveAs(filePath);
        }
    }
}