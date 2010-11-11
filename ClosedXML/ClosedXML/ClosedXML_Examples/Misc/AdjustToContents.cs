using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class AdjustToContents
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
            var ws = wb.Worksheets.Add("Adjust To Contents");

            // Set some values with different font sizes
            ws.Cell(2, 2).Value = "A";
            ws.Cell(2, 2).Style.Font.FontSize = 30;
            ws.Cell(3, 2).Value = "really, really, long text";
            ws.Cell(4, 2).Value = "long text";
            ws.Cell(5, 2).Value = "really long text";
            ws.Cell(5, 2).Style.Font.FontSize = 20;

            // Adjust all rows/columns in one shot
            ws.Rows().AdjustToContents();
            ws.Columns().AdjustToContents();

            // You can also adjust specific rows/columns

            // Adjust the width of column 2 to its contents
            //ws.Column(2).AdjustToContents();

            // Adjust the height of rows 2,3,4,5 to their contents
            //ws.Rows(2, 5).AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
