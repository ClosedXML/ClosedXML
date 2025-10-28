using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Rows;

internal class XLRowsTests
{
    [Test]
    public void Style_sets_format_of_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Rows("2,4-5").Style.Font.FontSize = 5;
        ws.Range("A1:A6").Style.Font.Bold = true; // Materialize formats in column A

        var expected = new[]
        {
            (Row: 1, FontSize: 11),
            (Row: 2, FontSize: 5),
            (Row: 3, FontSize: 11),
            (Row: 4, FontSize: 5),
            (Row: 5, FontSize: 5),
            (Row: 6, FontSize: 11),
        };
        foreach (var (row, fontSize) in expected)
        {
            Assert.AreEqual(fontSize, ws.Row(row).Style.Font.FontSize);

            var cellWithFormat = ws.Cell("A" + row);
            Assert.AreEqual(fontSize, cellWithFormat.Style.Font.FontSize);
            Assert.True(cellWithFormat.Style.Font.Bold);

            var nonMaterializedCell = ws.Cell("B" + row);
            Assert.AreEqual(fontSize, nonMaterializedCell.Style.Font.FontSize);
        }
    }
}
