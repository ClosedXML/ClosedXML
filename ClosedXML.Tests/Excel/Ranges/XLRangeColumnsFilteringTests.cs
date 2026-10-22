using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Ranges;

/// <summary>
/// Tests for filtering range columns (<see cref="IXLRangeColumns"/>).
/// </summary>
internal class XLRangeColumnsFilteringTests
{
    [Test]
    public void Skip_then_AdjustToContents_does_not_change_skipped_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var insertedRange = ws.Cell("A1").InsertData(new[]
        {
            new object[] { "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 }
        });
        insertedRange.FirstColumn().Style.Alignment.SetWrapText(true);
        insertedRange.FirstColumn().WorksheetColumn().Width = 20;
        ws.Column(2).Width = 1;

        IXLRangeColumns remaining = insertedRange.ColumnsUsed().Skip(1);
        remaining.AdjustToContents();

        Assert.AreEqual(20, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 1);
    }

    [Test]
    public void AdjustToContents_uses_only_cells_inside_the_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = "x";
        ws.Cell("B20").Value = "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!";
        ws.Column(2).Width = 1;

        ws.Range("B1:B1").FirstColumn().AdjustToContents();
        var widthFromRangeColumn = ws.Column(2).Width;

        ws.Range("A1:B1").ColumnsUsed().Skip(1).AdjustToContents();
        var widthFromRangeColumns = ws.Column(2).Width;

        ws.Column(2).AdjustToContents();
        Assert.Greater(ws.Column(2).Width, widthFromRangeColumn);
        Assert.Greater(ws.Column(2).Width, widthFromRangeColumns);
    }

    [Test]
    public void Skip_Where_Take_use_worksheet_column_numbers()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("C1:E1").Value = "x";

        IXLRangeColumns skipped = ws.Range("C1:E1").ColumnsUsed().Skip(1);
        IXLRangeColumns filtered = ws.Range("C1:E1").ColumnsUsed().Where(c => c.ColumnNumber() != 3);
        IXLRangeColumns taken = ws.Range("C1:E1").ColumnsUsed().Take(2);
        IXLRangeColumns chained = ws.Range("C1:E1").ColumnsUsed().Skip(1).Take(1);

        CollectionAssert.AreEqual(new[] { 4, 5 }, skipped.Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 4, 5 }, filtered.Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 3, 4 }, taken.Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 4 }, chained.Select(c => c.ColumnNumber()));
    }

    [Test]
    public void Skip_then_Delete_does_not_delete_first_range_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "keep";
        ws.Cell("B1").Value = "drop";
        ws.Cell("C1").Value = "drop";

        IXLRangeColumns remaining = ws.Range("A1:C1").ColumnsUsed().Skip(1);
        remaining.Delete();

        Assert.AreEqual("keep", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("B1").IsEmpty());
        Assert.True(ws.Cell("C1").IsEmpty());
    }

    [Test]
    public void ColumnsUsed_predicate_filters_range_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("C1:E1").Value = "x";

        CollectionAssert.AreEqual(new[] { 4, 5 }, ws.Range("C1:E1").ColumnsUsed().Skip(1).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 4, 5 }, ws.Range("C1:E1").ColumnsUsed(c => c.ColumnNumber() != 3).Select(c => c.ColumnNumber()));
    }
}
