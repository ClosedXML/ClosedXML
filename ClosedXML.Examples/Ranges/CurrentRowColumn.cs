using ClosedXML.Excel;

namespace ClosedXML.Examples.Ranges
{
    public class CurrentRowColumn : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Current Row Column");

            var cell = ws.Cell(5, 2);
            cell.Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(1, 1)
                .SetValue("Red's Row:")
                .CellRight().SetValue(cell.WorksheetRow().RowNumber())
                .CellBelow().SetValue(cell.WorksheetColumn().ColumnLetter())
                .CellLeft().SetValue("Red's Column:");

            var row = ws.Range("A6:C6").FirstRow();
            row.Style.Fill.SetBackgroundColor(XLColor.Blue);

            var column = ws.Range("B7:B9").FirstColumn();
            column.Style.Fill.SetBackgroundColor(XLColor.Green);

            ws.Cell(1, 4)
                .SetValue("Blue's Row:")
                .CellRight().SetValue(row.WorksheetRow().RowNumber())
                .CellBelow().SetValue(column.WorksheetColumn().ColumnLetter())
                .CellLeft().SetValue("Green's Column:");

            ws.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}