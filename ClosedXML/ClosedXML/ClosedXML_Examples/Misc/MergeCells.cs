using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class MergeCells
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

        #region Constructors

        // Public


        // Private


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
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Merge Cells");

            ws.Cell("B2").Value = "Merged Cells (B2 - D2)";
            ws.Range("B2:D2").Merge();

            ws.Cell("B4").Value = "Merged Cells (B4 - D6)";
            ws.Cell("B4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("B4:D6").Merge();

            ws.Cell("B8").Value = "Unmerged";
            ws.Range("B8:D8").Merge();
            ws.Range("B8:D8").Unmerge();

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
