using ClosedXML.Excel;

namespace ClosedXML.Examples.Ranges
{
    public class WalkingRanges : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Walking Cells");

            var cell = ws.Cell(5, 5).SetValue("(5,5)");

            cell.CellAbove().SetValue("(4,5)").Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            cell.CellAbove(2).SetValue("(3,5)").Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            cell.CellBelow().SetValue("(6,5)").Style.Fill.SetBackgroundColor(XLColor.Salmon);
            cell.CellBelow(2).SetValue("(7,5)").Style.Fill.SetBackgroundColor(XLColor.Salmon);

            cell.CellLeft().SetValue("(5,4)").Style.Fill.SetBackgroundColor(XLColor.LightBlue);
            cell.CellLeft(2).SetValue("(5,3)").Style.Fill.SetBackgroundColor(XLColor.LightBlue);
            cell.CellRight().SetValue("(5,6)").Style.Fill.SetBackgroundColor(XLColor.BlueBell);
            cell.CellRight(2).SetValue("(5,7)").Style.Fill.SetBackgroundColor(XLColor.BlueBell);

            var wsWalkRows = wb.Worksheets.Add("Walking rows");

            var row = wsWalkRows.Row(3);
            row.RowAbove().Style.Fill.SetBackgroundColor(XLColor.Salmon);
            row.RowAbove(2).Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            row.RowBelow().Style.Fill.SetBackgroundColor(XLColor.Blue);
            row.RowBelow(2).Style.Fill.SetBackgroundColor(XLColor.BlueBell);

            var rangeRow = wsWalkRows.Range("B8:D12").Row(3);
            rangeRow.RowAbove().Style.Fill.SetBackgroundColor(XLColor.Salmon);
            rangeRow.RowAbove(2).Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            rangeRow.RowBelow().Style.Fill.SetBackgroundColor(XLColor.Blue);
            rangeRow.RowBelow(2).Style.Fill.SetBackgroundColor(XLColor.BlueBell);

            var wsWalkColumns = wb.Worksheets.Add("Walking columns");

            var column = wsWalkColumns.Column(3);
            column.ColumnLeft().Style.Fill.SetBackgroundColor(XLColor.Salmon);
            column.ColumnLeft(2).Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            column.ColumnRight().Style.Fill.SetBackgroundColor(XLColor.Blue);
            column.ColumnRight(2).Style.Fill.SetBackgroundColor(XLColor.BlueBell);

            var rangeColumn = wsWalkColumns.Range("H2:L4").Column(3);
            rangeColumn.ColumnLeft().Style.Fill.SetBackgroundColor(XLColor.Salmon);
            rangeColumn.ColumnLeft(2).Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            rangeColumn.ColumnRight().Style.Fill.SetBackgroundColor(XLColor.Blue);
            rangeColumn.ColumnRight(2).Style.Fill.SetBackgroundColor(XLColor.BlueBell);

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}