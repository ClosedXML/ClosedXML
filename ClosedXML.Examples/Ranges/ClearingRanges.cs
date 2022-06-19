using ClosedXML.Excel;
using System.Linq;

namespace ClosedXML.Examples.Ranges
{
    public class ClearingRanges : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Clearing Ranges");
            foreach (var ro in Enumerable.Range(1, 10))
            {
                foreach (var co in Enumerable.Range(1, 10))
                {
                    var cell = ws.Cell(ro, co);
                    cell.Value = cell.Address.ToString();
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Fill.BackgroundColor = XLColor.Turquoise;
                    cell.Style.Font.Bold = true;
                }
            }

            // Clearing a range
            ws.Range("B1:C2").Clear();

            // Clearing a row in a range
            ws.Range("B4:C5").Row(1).Clear();

            // Clearing a column in a range
            ws.Range("E1:F4").Column(2).Clear();

            // Clear an entire row
            ws.Row(7).Clear();

            // Clear an entire column
            ws.Column("H").Clear();

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}