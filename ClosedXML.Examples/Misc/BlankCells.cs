using System;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Misc
{
    public class BlankCells : IXLExample
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
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).Value = "X";
            ws.Cell(1, 1).Clear();
            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
