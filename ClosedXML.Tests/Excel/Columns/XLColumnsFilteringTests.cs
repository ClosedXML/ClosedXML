using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Columns;

/// <summary>
/// Tests for filtering worksheet columns (<see cref="IXLColumns"/>).
/// </summary>
internal class XLColumnsFilteringTests
{
    [Test]
    public void Skip_then_AdjustToContents_does_not_change_skipped_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Cool cheesecake stuff");
        var insertedRange = ws.Cell("A1").InsertData(new[]
        {
            new object[] { "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 }
        });
        insertedRange.FirstColumn().Style.Alignment.SetWrapText(true);
        insertedRange.FirstColumn().WorksheetColumn().Width = 20;
        ws.Column(2).Width = 1;

        IXLColumns remaining = ws.ColumnsUsed().Skip(1);
        remaining.AdjustToContents();

        Assert.AreEqual(20, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 1);
    }

    [Test]
    public void Where_and_Take_then_AdjustToContents_filter_as_specified()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            new object[] { "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 }
        });
        ws.Column(1).Width = 20;
        ws.Column(2).Width = 1;

        IXLColumns remaining = ws.ColumnsUsed().Where(c => c.ColumnNumber() != 1);
        remaining.AdjustToContents();
        Assert.AreEqual(20, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 1);

        ws.Column(1).Width = 1;
        ws.Column(2).Width = 1;
        IXLColumns taken = ws.ColumnsUsed().Take(1);
        taken.AdjustToContents();
        Assert.Greater(ws.Column(1).Width, 1);
        Assert.AreEqual(1, ws.Column(2).Width, XLHelper.Epsilon);
    }

    [Test]
    public void Skip_Where_Take_enumerate_in_column_number_order()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").Value = "c";
        ws.Cell("A1").Value = "a";
        ws.Cell("B1").Value = "b";

        IXLColumns skipped = ws.ColumnsUsed().Skip(1);
        IXLColumns filtered = ws.ColumnsUsed().Where(c => c.ColumnNumber() != 2);
        IXLColumns taken = ws.ColumnsUsed().Take(2);

        CollectionAssert.AreEqual(new[] { 2, 3 }, skipped.Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 1, 3 }, filtered.Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 1, 2 }, taken.Select(c => c.ColumnNumber()));
        Assert.AreEqual(0, ws.ColumnsUsed().Skip(10).Count());
    }

    [Test]
    public void Filtering_Columns_does_not_treat_result_as_all_columns_of_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:C1").Value = "x";
        var defaultColumnWidth = ws.ColumnWidth;

        ws.Columns().Skip(1).Width = 100;
        Assert.AreEqual(defaultColumnWidth, ws.ColumnWidth, XLHelper.Epsilon);
        Assert.AreEqual(defaultColumnWidth, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.AreEqual(100, ws.Column(2).Width, XLHelper.Epsilon);
        Assert.AreEqual(100, ws.Column(3).Width, XLHelper.Epsilon);

        ws.Columns().Where(c => c.ColumnNumber() != 1).Width = 90;
        Assert.AreEqual(defaultColumnWidth, ws.ColumnWidth, XLHelper.Epsilon);
        Assert.AreEqual(defaultColumnWidth, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.AreEqual(90, ws.Column(2).Width, XLHelper.Epsilon);

        ws.Columns().Skip(0).Width = 80;
        Assert.AreEqual(defaultColumnWidth, ws.ColumnWidth, XLHelper.Epsilon);
        Assert.AreEqual(80, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.AreEqual(defaultColumnWidth, ws.Column("Z").Width, XLHelper.Epsilon);
    }

    [Test]
    public void Skip_on_Columns_Delete_does_not_clear_the_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "keep";
        ws.Cell("B1").Value = "drop";
        ws.Cell("C1").Value = "drop";

        IXLColumns skipped = ws.Columns().Skip(1);
        skipped.Delete();

        Assert.AreEqual("keep", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("B1").IsEmpty());
        Assert.True(ws.Cell("C1").IsEmpty());
    }

    [Test]
    public void Skip_Where_Take_can_be_chained()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:E1").Value = "x";

        IXLColumns subset = ws.ColumnsUsed().Skip(1).Where(c => c.ColumnNumber() != 4).Take(2);
        IXLColumns skippedTwice = ws.ColumnsUsed().Skip(1).Skip(1);

        CollectionAssert.AreEqual(new[] { 2, 3 }, subset.Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 3, 4, 5 }, skippedTwice.Select(c => c.ColumnNumber()));
    }

    [Test]
    public void AdjustToContents_overloads_respect_row_bounds_and_min_max()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = "x";
        ws.Cell("B10").Value = "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!";
        ws.Column(2).Width = 1;

        IXLColumns columnB = ws.ColumnsUsed().Where(c => c.ColumnNumber() == 2);
        columnB.AdjustToContents(1, 1);
        var widthFromFirstRow = ws.Column(2).Width;

        ws.Column(2).AdjustToContents();
        Assert.Greater(ws.Column(2).Width, widthFromFirstRow);

        ws.Column(2).Width = 1;
        columnB.AdjustToContents(50.0, 50.0);
        Assert.AreEqual(50, ws.Column(2).Width, XLHelper.Epsilon);
    }

    [Test]
    public void Collection_operations_do_not_affect_skipped_column_or_other_sheets()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("One");
        var ws2 = wb.AddWorksheet("Two");
        ws1.Cell("A1").Value = "keep";
        ws1.Cell("B1").Value = "change";
        ws1.Cell("C1").Value = "change";
        ws2.Cells("A1:C1").Value = "y";
        var ws2Width = ws2.Column(2).Width;

        IXLColumns remaining = ws1.ColumnsUsed().Skip(1);
        remaining.Style.Font.Bold = true;
        remaining.Hide();
        remaining.Clear(XLClearOptions.Contents);
        remaining.Width = 100;

        Assert.AreEqual("keep", ws1.Cell("A1").GetString());
        Assert.False(ws1.Column(1).Style.Font.Bold);
        Assert.False(ws1.Column(1).IsHidden);
        Assert.True(ws1.Cell("B1").IsEmpty());
        Assert.True(ws1.Column(2).Style.Font.Bold);
        Assert.True(ws1.Column(2).IsHidden);
        Assert.AreEqual(100, ws1.Column(2).Width, XLHelper.Epsilon);
        Assert.AreEqual(ws2Width, ws2.Column(2).Width, XLHelper.Epsilon);
    }

    [Test]
    public void Skip_works_on_materialized_and_inserted_column_collections()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 9;

        IXLColumns disjoint = ws.Columns("A,C,E").Skip(1);
        disjoint.Width = 12;
        Assert.AreNotEqual(12, ws.Column(1).Width);
        Assert.AreEqual(12, ws.Column(3).Width, XLHelper.Epsilon);
        Assert.AreEqual(12, ws.Column(5).Width, XLHelper.Epsilon);
        Assert.AreNotEqual(12, ws.ColumnWidth);

        IXLColumns inserted = ws.Column(1).InsertColumnsAfter(3);
        IXLColumns skippedInsert = inserted.Skip(1);
        skippedInsert.Width = 15;
        Assert.AreEqual(9, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.AreEqual(9, ws.Column(2).Width, XLHelper.Epsilon);
        Assert.AreEqual(15, ws.Column(3).Width, XLHelper.Epsilon);
        Assert.AreEqual(15, ws.Column(4).Width, XLHelper.Epsilon);
    }

    [Test]
    public void Foreach_Skip_adjusts_only_remaining_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            new object[] { "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 }
        });
        ws.Column(1).Width = 20;
        ws.Column(2).Width = 1;

        foreach (var column in ws.ColumnsUsed().Skip(1))
            column.AdjustToContents();

        Assert.AreEqual(20, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 1);
    }

    [Test]
    public void ColumnsUsed_predicate_filters_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            new object[] { "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 }
        });
        ws.Column(1).Width = 20;
        ws.Column(2).Width = 1;

        ws.ColumnsUsed(c => c.ColumnNumber() != 1).AdjustToContents();

        Assert.AreEqual(20, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 1);
    }

    [Test]
    public void Skip_follows_column_number_order_not_insertion_order()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").Value = "c";
        ws.Cell("A1").Value = "a";
        ws.Cell("B1").Value = "b";

        CollectionAssert.AreEqual(new[] { 2, 3 }, ws.ColumnsUsed().Skip(1).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 3, 5 }, ws.Columns("A,C,E").Skip(1).Select(c => c.ColumnNumber()));
    }

    [Test]
    public void Indexed_Where_filters_by_index()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:C1").Value = "x";

        CollectionAssert.AreEqual(new[] { 2, 3 }, ws.ColumnsUsed().Where((c, i) => i > 0).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 2, 1 }, ws.ColumnsUsed().OrderByDescending(c => c.ColumnNumber()).Skip(1).Select(c => c.ColumnNumber()));
    }

    [Test]
    public void FindColumns_returns_matching_columns()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("One");
        var ws2 = wb.AddWorksheet("Two");
        ws1.Cells("A1:C1").Value = "x";
        ws2.Cells("B1:D1").Value = "y";

        var found = wb.FindColumns(c => c.ColumnNumber() == 2);

        Assert.AreEqual(2, found.Count());
        CollectionAssert.AreEquivalent(new[] { "One", "Two" }, found.Select(c => c.Worksheet.Name));
    }

    [Test]
    public void Columns_Width_and_Delete_affect_whole_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:C1").Value = "x";

        ws.Columns().Width = 100;
        Assert.AreEqual(100, ws.ColumnWidth, XLHelper.Epsilon);
        Assert.AreEqual(100, ws.Column("G").Width, XLHelper.Epsilon);

        ws.Columns().Delete();
        Assert.True(ws.Cell("A1").IsEmpty());
        Assert.True(ws.Cell("B1").IsEmpty());
        Assert.True(ws.Cell("C1").IsEmpty());
    }
}
