using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Ranges;

/// <summary>
/// Tests for filtering range rows (<see cref="IXLRangeRows"/>).
/// </summary>
internal class XLRangeRowsFilteringTests
{
    [Test]
    public void Skip_then_AdjustToContents_does_not_change_skipped_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Keep this height";
        ws.Cell("A2").Value = "Fit me";
        ws.Row(1).Height = 40;
        ws.Row(2).Height = 5;

        IXLRangeRows remaining = ws.Range("A1:A2").RowsUsed().Skip(1);
        remaining.AdjustToContents();

        Assert.AreEqual(40, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.AreNotEqual(5, ws.Row(2).Height);
    }

    [Test]
    public void AdjustToContents_uses_only_cells_inside_the_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A2").Value = "short";
        ws.Cell("A2").Style.Font.FontSize = 11;
        ws.Cell("Z2").Value = "tall";
        ws.Cell("Z2").Style.Font.FontSize = 36;
        ws.Row(2).Height = 10;

        ws.Range("A2:B2").FirstRow().AdjustToContents();
        var heightFromRangeRow = ws.Row(2).Height;
        Assert.Greater(heightFromRangeRow, 10);

        ws.Range("A2:B2").RowsUsed().AdjustToContents();
        var heightFromRangeRows = ws.Row(2).Height;

        ws.Row(2).AdjustToContents();
        Assert.Greater(ws.Row(2).Height, heightFromRangeRow);
        Assert.Greater(ws.Row(2).Height, heightFromRangeRows);
    }

    [Test]
    public void Skip_Where_Take_use_worksheet_row_numbers()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A3:A5").Value = "x";

        IXLRangeRows skipped = ws.Range("A3:A5").RowsUsed().Skip(1);
        IXLRangeRows filtered = ws.Range("A3:A5").RowsUsed().Where(r => r.RowNumber() != 3);
        IXLRangeRows taken = ws.Range("A3:A5").RowsUsed().Take(2);

        CollectionAssert.AreEqual(new[] { 4, 5 }, skipped.Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 4, 5 }, filtered.Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 3, 4 }, taken.Select(r => r.RowNumber()));
    }

    [Test]
    public void Skip_then_Delete_does_not_delete_first_range_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "keep";
        ws.Cell("A2").Value = "drop";
        ws.Cell("A3").Value = "drop";

        IXLRangeRows remaining = ws.Range("A1:A3").RowsUsed().Skip(1);
        remaining.Delete();

        Assert.AreEqual("keep", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("A2").IsEmpty());
        Assert.True(ws.Cell("A3").IsEmpty());
    }

    [Test]
    public void RowsUsed_predicate_filters_range_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A3:A5").Value = "y";

        CollectionAssert.AreEqual(new[] { 4, 5 }, ws.Range("A3:A5").RowsUsed().Skip(1).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 4, 5 }, ws.Range("A3:A5").RowsUsed(r => r.RowNumber() != 3).Select(r => r.RowNumber()));
    }
}
