using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;
using ClosedXML;

namespace ClosedXML_Examples
{
    public class LambdaExpressions
    {
        public void Create()
        {
            var workbook = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var ws = workbook.Worksheets.Worksheet(0);

            // Define a range with the data
            var firstDataCell = ws.Cell("B4");
            var lastDataCell = ws.LastCellUsed();
            var rngData = ws.Range(firstDataCell.Address, lastDataCell.Address);

            // Delete all rows where Outcast = false (the 3rd column)
            rngData.Rows() // From all rows
                .Where(r => !r.Cell(3).GetBoolean()) // where the 3rd cell of each row is false
                .ForEach(r => r.Delete()); // delete the row and shift the cells up (the default for rows in a range)

            // Put a light gray background to all text cells
            rngData.Cells() // From all cells
                .Where(c => c.DataType == XLCellValues.Text) // where the data type is Text
                .ForEach(c => c.Style.Fill.BackgroundColor = Color.LightGray); // Fill with a light gray

            // Put a thick border to the bottom of the table (we may have deleted the bottom cells with the border)
            rngData.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            workbook.SaveAs(@"C:\Excel Files\Created\LambdaExpressions.xlsx");
        }
    }
}