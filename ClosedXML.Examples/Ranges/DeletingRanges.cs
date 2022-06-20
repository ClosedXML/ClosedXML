using ClosedXML.Excel;

namespace ClosedXML.Examples.Ranges
{
    public class DeletingRanges : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Deleting Ranges");

            // Deleting Columns
            // Setup test values
            ws.Columns("1-3, 5, 7").Style.Fill.BackgroundColor = XLColor.Gray;
            ws.Columns("4, 6").Style.Fill.BackgroundColor = XLColor.GreenPigment;
            ws.Row(1).Cells("1-3, 5, 7").Value = "FAIL";

            ws.Column(7).Delete();
            ws.Column(1).Delete();
            ws.Columns(1, 2).Delete();
            ws.Column(2).Delete();

            // Deleting Rows
            ws.Rows("1,5,7").Style.Fill.BackgroundColor = XLColor.GreenPigment;
            ws.Rows("2-4,6, 8").Style.Fill.BackgroundColor = XLColor.Gray;
            ws.Column(1).Cells("2-4,6, 8").Value = "FAIL";

            ws.Row(8).Delete();
            ws.Row(2).Delete();
            ws.Rows(2, 3).Delete();
            ws.Rows(3, 4).Delete();

            // Deleting Ranges (Shifting Left)
            var rng1 = ws.Range(2, 2, 8, 8);
            rng1.Columns("1-3, 5, 7").Style.Fill.BackgroundColor = XLColor.Gray;
            rng1.Columns("4, 6").Style.Fill.BackgroundColor = XLColor.Orange;
            rng1.Row(1).Cells("1-3, 5, 7").Value = "FAIL";

            rng1.Column(7).Delete();
            rng1.Column(1).Delete();
            rng1.Range(1, 1, rng1.RowCount(), 2).Delete(XLShiftDeletedCells.ShiftCellsLeft);
            rng1.Column(2).Delete();

            // Deleting Ranges (Shifting Up)
            rng1.Rows("4, 6").Style.Fill.BackgroundColor = XLColor.Orange;
            rng1.Rows("1-3, 5, 7").Style.Fill.BackgroundColor = XLColor.Gray;
            rng1.Column(1).Cells("1-3, 5, 7").Value = "FAIL";

            rng1.Row(7).Delete();
            rng1.Row(1).Delete();
            rng1.Range(1, 1, 2, rng1.ColumnCount()).Delete(XLShiftDeletedCells.ShiftCellsUp);
            rng1.Row(2).Delete();

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}