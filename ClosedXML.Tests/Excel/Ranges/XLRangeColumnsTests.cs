using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges;

internal class XLRangeColumnsTests
{
    [Test]
    public void Style_sets_format_of_range_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("B3:E4");
        var rangeColumns = range.Columns("A,C-D");

        rangeColumns.Style.Font.FontSize = 20;

        var expectedChangedCells = new[] { "B3", "B4", "D3", "D4", "E3", "E4" }.ToHashSet();
        foreach (var cell in range.Grow().Cells())
        {
            var address = cell.Address.ToString();
            var fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(fontSize, cell.Style.Font.FontSize);
        }
    }
}
