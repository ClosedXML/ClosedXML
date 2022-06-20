using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class HideSheets : IXLExample
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

            wb.Worksheets.Add("First Hidden").Hide();
            wb.Worksheets.Add("Visible");
            wb.Worksheets.Add("Unhidden").Hide().Unhide();
            wb.Worksheets.Add("VeryHidden").Visibility = XLWorksheetVisibility.VeryHidden;
            wb.Worksheets.Add("Last Hidden").Hide();

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}