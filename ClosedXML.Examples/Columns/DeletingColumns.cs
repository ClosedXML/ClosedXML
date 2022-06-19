using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class DeletingColumns : IXLExample
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
            var ws = workbook.Worksheets.Add("Deleting Columns");

            var rngTitles = ws.Range("B2:D2");
            ws.Row(1).InsertRowsBelow(2);

            var rng1 = ws.Range("B2:D2");
            var rng2 = ws.Range("F2:G2");
            var rng3 = ws.Range("A1:A3");
            var col1 = ws.Column(1);

            rng1.Style.Fill.BackgroundColor = XLColor.Orange;
            rng2.Style.Fill.BackgroundColor = XLColor.Blue;
            rng3.Style.Fill.BackgroundColor = XLColor.Red;
            col1.Style.Fill.BackgroundColor = XLColor.Black;

            ws.Columns("A,C,E:H").Delete();
            ws.Cell("A2").Value = "OK";
            ws.Cell("B2").Value = "OK";

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}