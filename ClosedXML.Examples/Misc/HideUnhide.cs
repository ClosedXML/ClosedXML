using System;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Misc
{
    public class HideUnhide : IXLExample
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Hide Rows Columns");

            ws.Columns(1, 3).Hide();
            ws.Rows(1, 3).Hide();

            ws.Column(2).Unhide();
            ws.Row(2).Unhide();

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
