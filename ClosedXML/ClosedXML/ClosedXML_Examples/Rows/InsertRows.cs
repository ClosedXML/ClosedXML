using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Examples.Rows
{
    public class InsertRows
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
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Insert Rows");

            foreach (var r in Enumerable.Range(1, 5))
            {
                foreach (var c in Enumerable.Range(1, 5))
                {
                    ws.Cell(r, c).Value = "X";
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
            }

            ws.Row(3).InsertRowsBelow(2);
            ws.Row(1).InsertRowsAbove(2);
            ws.Range("D3:E4").InsertRowsBelow(2);
            ws.Range("B3:C5").InsertRowsAbove(2);

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
