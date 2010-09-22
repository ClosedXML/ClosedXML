using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Ranges
{
    public class DeletingRanges
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();

            var ws = workbook.Worksheets.Add("Deleting Ranges");
            foreach (var ro in Enumerable.Range(1, 10))
                foreach (var co in Enumerable.Range(1, 10))
                    ws.Cell(ro, co).Value = ws.Cell(ro, co).Address.ToString();
            
            // Delete range and shift cells up
            ws.Range("B4:C5").Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Delete range and shift cells left
            ws.Range("D1:E3").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            // Delete an entire row
            ws.Row(5).Delete();

            // Delete a row in a range, shift cells up
            ws.Range("A1:C4").Row(2).Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Delete an entire column
            ws.Column(5).Delete();

            // Delete a column in a range, shift cells up
            ws.Range("A1:C4").Column(2).Delete(XLShiftDeletedCells.ShiftCellsLeft);

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
