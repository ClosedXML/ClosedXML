using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Tables;

internal class XLTableRowsTests
{
    [Test]
    public void Style_sets_format_of_rows()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var table = ws.Cell("B3").InsertTable(new[]
        {
            new { Name = "Cake", Price = 7},
            new { Name = "Waffle", Price = 4},
            new { Name = "Croissant", Price = 5},
            new { Name = "Pie", Price = 9 },
        });
        var rows = table.Rows(x => x.Cell(1).GetText().StartsWith("C"));

        // Act
        rows.Style.Font.FontSize = 20;

        // Assert
        var expectedChangedCells = new[]
        {
            "B4", "C4", // Cake row
            "B6", "C6", // Croissant row
        }.ToHashSet();

        foreach (var cell in ws.Range("A2:D7").Cells())
        {
            var address = cell.Address.ToString();
            var expectedFontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(expectedFontSize, cell.Style.Font.FontSize);
        }
    }
}
