using System.Drawing;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Performance
{
    [MemoryDiagnoser]
    public class WorkbookOperationBenchmarks
    {
        const int rowCount = 100;
        const int colCount = 11;

        [Benchmark]
        public void CellFillAndCopy()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            // Create a source column with formatting
            var sourceColumn = ws.Column(1);
            FormatColumn(sourceColumn);

            // Create formatted cells in columns 2-11, rows 1-100
            for (var col = 2; col <= colCount; col++)
            {
                for (var row = 1; row <= rowCount; row++) FormatCell(ws, row, col);
            }

            // Now copy the source column to each of these cells
            for (var col = 2; col <= colCount; col++)
            {
                for (var row = 1; row <= rowCount; row++)
                {
                    // Copy the entire source column to this cell
                    ws.Cell(row, col).CopyFrom(sourceColumn);
                }
            }
        }

        private static void FormatCell(IXLWorksheet ws, int row, int col)
        {
            ws.Cell(row, col).Style.Fill.SetBackgroundColor(XLColor.FromArgb(row % 255, col % 255, (row + col) % 255));
            ws.Cell(row, col).Style.Font.Bold = row % 2 == 0;
            ws.Cell(row, col).Style.Font.Italic = col % 2 == 0;
            ws.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private static void FormatColumn(IXLColumn sourceColumn)
        {
            sourceColumn.Cell(1).Style.Fill.SetBackgroundColor(XLColor.Red);
            sourceColumn.Cell(2).Style.Fill.SetBackgroundColor(XLColor.FromArgb(1, 1, 1));
            sourceColumn.Cell(3).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#CCCCCC"));
            sourceColumn.Cell(4).Style.Fill.SetBackgroundColor(XLColor.FromIndex(26));
            sourceColumn.Cell(5).Style.Fill.SetBackgroundColor(XLColor.FromColor(Color.MediumSeaGreen));
            sourceColumn.Cell(6).Style.Fill.SetBackgroundColor(XLColor.FromName("Blue"));
            sourceColumn.Cell(7).Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent3));
        }
    }
}