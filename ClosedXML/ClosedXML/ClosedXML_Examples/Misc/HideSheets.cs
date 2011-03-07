using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class HideSheets
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

            wb.Worksheets.Add("Visible");
            wb.Worksheets.Add("Hidden").Hide();
            wb.Worksheets.Add("Unhidden").Hide().Unhide();
            wb.Worksheets.Add("VeryHidden").Visibility = XLWorksheetVisibility.VeryHidden;

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
