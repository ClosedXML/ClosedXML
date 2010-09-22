using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Ranges
{
    public class DefiningRanges
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Defining a Range");

            // With a string
            var range1 = ws.Range("A1:B1");
            range1.Cell(1, 1).Value = "ws.Range(\"A1:B1\").Merge()";
            range1.Merge();

            // With two XLAddresses
            var range2 = ws.Range(ws.Cell(2, 1).Address, ws.Cell(2, 2).Address);
            range2.Cell(1, 1).Value = "ws.Range(ws.Cell(2, 1).Address, ws.Cell(2, 2).Address).Merge()";
            range2.Merge();

            // With two XLCells
            var range3 = ws.Range(ws.Cell(3,1), ws.Cell(3,2));
            range3.Cell(1, 1).Value = "ws.Range(ws.Cell(3,1), ws.Cell(3,2)).Merge()";
            range3.Merge();

            // With two strings
            var range4 = ws.Range("A4", "B4");
            range4.Cell(1, 1).Value = "ws.Range(\"A4\", \"B4\").Merge()";
            range4.Merge();

            // With 4 points
            var range5 = ws.Range(5, 1, 5, 2);
            range5.Cell(1, 1).Value = "ws.Range(5, 1, 5, 2).Merge()";
            range5.Merge();

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
