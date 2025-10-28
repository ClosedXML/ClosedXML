using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges;

internal class XLRangeRowsTests
{
    [Test]
    public void Style_sets_format_of_range_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("B3:C7");
        var rangeRows = range.Rows("1,3-4");

        rangeRows.Style.Font.FontSize = 20;

        var expectedChangedCells = new[] { "B3", "C3", "B5", "C5", "B6", "C6" }.ToHashSet();
        foreach (var cell in range.Grow().Cells())
        {
            var address = cell.Address.ToString();
            var fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(fontSize, cell.Style.Font.FontSize, 0 , address);
        }
    }
}
