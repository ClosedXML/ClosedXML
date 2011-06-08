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
            ws.Cell(1, 1).Value = "Tall Row";
            ws.Cell(1, 1).Style.Font.FontSize = 30;
            ws.Cell(2, 1).Value = "Very Wide Column";
            ws.Cell(2, 1).Style.Font.FontSize = 20;
            
            // Adjust column width
            ws.Column(1).AdjustToContents();

            // Adjust row heights
            ws.Rows(1, 2).AdjustToContents();

            // You can also adjust all rows/columns in one shot
            // ws.Rows().AdjustToContents();
            // ws.Columns().AdjustToContents();


            // We'll now select which cells should be used for calculating the
            // column widths (same method applies for row heights)
            
            // Set the values
            ws.Cell(4, 2).Value = "Width ignored because calling column.AdjustToContents(5, 7)";
            ws.Cell(5, 2).Value = "Short text";
            ws.Cell(6, 2).Value = "Width ignored because it's part of a merge";
            ws.Range(6, 2, 6, 4).Merge();
            ws.Cell(7, 2).Value = "Width should adjust to this cell";
            ws.Cell(8, 2).Value = "Width ignored because calling column.AdjustToContents(5, 7)";

            // Adjust column widths only taking into account rows 5-7
            // (merged cells will be ignored)
            ws.Column(2).AdjustToContents(5, 7);

            // You can also specify the starting row to start calculating the widths:
            // e.g. ws.Column(3).AdjustToContents(9);

            var ws2 = wb.Worksheets.Add("Adjust Widths");
            ws2.Cell(1, 1).SetValue("Text to adjust - 255").Style.Alignment.TextRotation = 255;
            for (Int32 co = 0; co < 90; co += 5)
            {
                ws2.Cell(1, (co / 5) + 2).SetValue("Text to adjust - " + co).Style.Alignment.TextRotation = co;
            }

            ws2.Columns().AdjustToContents();

            var ws3 = wb.Worksheets.Add("Adjust Heights");
            ws3.Cell(1, 1).SetValue("Text to adjust - 255").Style.Alignment.TextRotation = 255;
            for (Int32 ro = 0; ro < 90; ro += 5)
            {
                ws3.Cell((ro / 5) + 2, 1).SetValue("Text to adjust - " + ro).Style.Alignment.TextRotation = ro;
            }

            ws3.Rows().AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
