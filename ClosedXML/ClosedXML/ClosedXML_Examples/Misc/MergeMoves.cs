using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class MergeMoves
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
        public void Create()
        {
            var workbook = new XLWorkbook(@"C:\Excel Files\Created\MergedCells.xlsx");
            var ws = workbook.Worksheet(1);

            ws.Range("B1:F1").InsertRowsBelow(1);
            ws.Range("A3:A9").InsertColumnsAfter(1);
            ws.Row(1).Delete();
            ws.Column(1).Delete();

            ws.Range("E8:E9").InsertColumnsAfter(1);
            ws.Range("F2:F8").Merge();
            ws.Range("E3:E4").InsertColumnsAfter(1);
            ws.Range("F2:F8").Merge();
            ws.Range("E1:E2").InsertColumnsAfter(1);
            ws.Range("G2:G8").Merge();
            ws.Range("E1:E2").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            ws.Range("D3:E3").InsertRowsBelow(1);
            ws.Range("A1:B1").InsertRowsBelow(1);
            ws.Range("B3:D3").Merge();
            ws.Range("A1:B1").Delete(XLShiftDeletedCells.ShiftCellsUp);

            ws.Range("B8:D8").Merge();
            ws.Range("D8:D9").Clear();

            workbook.SaveAs(@"C:\Excel Files\Created\MergedMoves.xlsx");
        }

        // Private

        // Override


        #endregion
    }
}
