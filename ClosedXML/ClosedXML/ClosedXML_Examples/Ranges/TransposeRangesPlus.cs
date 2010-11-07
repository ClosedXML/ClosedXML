using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;


namespace ClosedXML_Examples
{
    public class TransposeRangesPlus
    {
        public void Create()
        {
            var workbook = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var ws = workbook.Worksheets.GetWorksheet(0);

            var rngTable = ws.Range("B2:F6");

            rngTable.Row(rngTable.RowCount() - 1).Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Place some markers
            var cellNextRow = ws.Cell(rngTable.LastAddressInSheet.RowNumber + 1, rngTable.LastAddressInSheet.ColumnNumber);
            cellNextRow.Value = "Next Row";
            var cellNextColumn = ws.Cell(rngTable.LastAddressInSheet.RowNumber, rngTable.LastAddressInSheet.ColumnNumber + 1);
            cellNextColumn.Value = "Next Column";

            rngTable.Transpose(XLTransposeOptions.MoveCells);
            rngTable.Transpose(XLTransposeOptions.MoveCells);
            rngTable.Transpose(XLTransposeOptions.ReplaceCells);
            rngTable.Transpose(XLTransposeOptions.ReplaceCells);

            workbook.SaveAs(@"C:\Excel Files\Created\TransposeRangesPlus.xlsx");
        }
    }
}