using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Rows;

/// <summary>
/// Tests for filtering worksheet rows (<see cref="IXLRows"/>).
/// </summary>
internal class XLRowsFilteringTests
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

        IXLRows remaining = ws.RowsUsed().Skip(1);
        remaining.AdjustToContents();

        Assert.AreEqual(40, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.AreNotEqual(5, ws.Row(2).Height);
    }

    [Test]
    public void Filtering_Rows_does_not_treat_result_as_all_rows_of_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "keep";
        ws.Cell("A2").Value = "drop";
        ws.Cell("A3").Value = "drop";
        var defaultRowHeight = ws.RowHeight;

        ws.Rows().Skip(1).Height = 30;
        Assert.AreEqual(defaultRowHeight, ws.RowHeight, XLHelper.Epsilon);
        Assert.AreEqual(defaultRowHeight, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.AreEqual(30, ws.Row(2).Height, XLHelper.Epsilon);
        Assert.AreEqual(30, ws.Row(3).Height, XLHelper.Epsilon);
        Assert.AreEqual(defaultRowHeight, ws.Row(11).Height, XLHelper.Epsilon);

        ws.Rows().Skip(0).Height = 25;
        Assert.AreEqual(defaultRowHeight, ws.RowHeight, XLHelper.Epsilon);

        ws.Rows().Skip(1).Delete();
        Assert.AreEqual("keep", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("A2").IsEmpty());
        Assert.True(ws.Cell("A3").IsEmpty());
    }

    [Test]
    public void Skip_Where_Take_can_be_chained()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:A5").Value = "x";

        IXLRows subset = ws.RowsUsed().Skip(1).Where(r => r.RowNumber() != 4).Take(2);

        CollectionAssert.AreEqual(new[] { 2, 3 }, subset.Select(r => r.RowNumber()));
    }

    [Test]
    public void Foreach_Skip_adjusts_only_remaining_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Keep this height";
        ws.Cell("A2").Value = "Fit me";
        ws.Row(1).Height = 40;
        ws.Row(2).Height = 5;

        foreach (var row in ws.RowsUsed().Skip(1))
            row.AdjustToContents();

        Assert.AreEqual(40, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.AreNotEqual(5, ws.Row(2).Height);
    }

    [Test]
    public void RowsUsed_predicate_filters_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Keep this height";
        ws.Cell("A2").Value = "Fit me";
        ws.Row(1).Height = 40;
        ws.Row(2).Height = 5;

        ws.RowsUsed(r => r.RowNumber() != 1).AdjustToContents();

        Assert.AreEqual(40, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.AreNotEqual(5, ws.Row(2).Height);
    }

    [Test]
    public void Rows_Height_and_Delete_affect_whole_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:A3").Value = "x";

        ws.Rows().Height = 30;
        Assert.AreEqual(30, ws.RowHeight, XLHelper.Epsilon);
        Assert.AreEqual(30, ws.Row(11).Height, XLHelper.Epsilon);

        ws.Rows().Delete();
        Assert.True(ws.Cell("A1").IsEmpty());
        Assert.True(ws.Cell("A2").IsEmpty());
        Assert.True(ws.Cell("A3").IsEmpty());
    }
}
