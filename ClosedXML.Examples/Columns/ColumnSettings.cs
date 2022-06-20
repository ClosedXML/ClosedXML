using ClosedXML.Excel;

namespace ClosedXML.Examples.Columns
{
    public class ColumnSettings : IXLExample
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
        public ColumnSettings()
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
            var ws = workbook.Worksheets.Add("Column Settings");

            var col1 = ws.Column("B");
            col1.Style.Fill.BackgroundColor = XLColor.Red;
            col1.Width = 20;

            var col2 = ws.Column(4);
            col2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
            col2.Width = 5;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}