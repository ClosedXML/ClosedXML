using ClosedXML.Excel;

namespace ClosedXML.Examples.Rows
{
    public class RowSettings : IXLExample
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

        #region Constructors

        // Public
        public RowSettings()
        {
        }

        // Private

        #endregion Constructors

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
            var ws = workbook.Worksheets.Add("Row Settings");

            var row1 = ws.Row(2);
            row1.Style.Fill.BackgroundColor = XLColor.Red;
            row1.Height = 30;

            var row2 = ws.Row(4);
            row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
            row2.Height = 3;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}