using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class RowCells : IXLExample
    {
        #region Variables

        // Public

        // Private

        #endregion Variables

        #region Properties

        // Public

        // Private

        // Override

        #endregion Properties

        #region Events

        // Public

        // Private

        // Override

        #endregion Events

        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Row Cells");

            var rowFromWorksheet = ws.Row(1);
            rowFromWorksheet.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
            rowFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
            rowFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
            rowFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

            var rowFromRange = ws.Range("A2:I2").FirstRow();

            rowFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
            rowFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
            rowFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
            rowFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

            ws.Columns().Width = 7;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}