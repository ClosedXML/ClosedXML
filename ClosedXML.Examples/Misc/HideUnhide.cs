using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class HideUnhide : IXLExample
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
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Hide Rows Columns");

            ws.Columns(1, 3).Hide();
            ws.Rows(1, 3).Hide();

            ws.Column(2).Unhide();
            ws.Row(2).Unhide();

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}