using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Hyperlinks;

[TestFixture]
[TestOf(typeof(XLHyperlinks))]
public class XLHyperlinksTests
{
    [TestCaseSource(nameof(StructuralChangeCases))]
    public void Hyperlink_is_moved_on_sheet_structure_change(string hyperlinkPosition, Action<IXLWorksheet> structuralChange, string expectedPosition)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var hyperlink = new XLHyperlink("https://example.com");
        ws.Cell(hyperlinkPosition).SetHyperlink(hyperlink);

        structuralChange(ws);

        Assert.False(ws.Cell(hyperlinkPosition).HasHyperlink);
        Assert.AreSame(ws.Cell(expectedPosition).GetHyperlink(), hyperlink);
    }

    public static IEnumerable<object[]> StructuralChangeCases
    {
        get
        {
            return new List<(string, Action<IXLWorksheet>, string)>
            {
                ("D5", ws => ws.Range("A5:B5").Delete(XLShiftDeletedCells.ShiftCellsLeft), "B5"),
                ("D5", ws => ws.Range("B2:D4").Delete(XLShiftDeletedCells.ShiftCellsUp), "D2"),
                ("D5", ws => ws.Column("D").InsertColumnsBefore(2), "F5"), // Insert column leftward
                ("D5", ws => ws.Row(2).InsertRowsAbove(4), "D9"), // Insert row above
            }.Select(x => new object[] { x.Item1, x.Item2, x.Item3 });
        }
    }

    [Test]
    public void Shift_doesnt_collide_hyperlinks()
    {
        // In former original data structures, there could be only one hyperlink per area
        // and when links were shifted, one link could shift to a position of another that
        // hasn't been yet shifted. New data structure allows multiple links in a same area,
        // though I hope it's a rare occurence.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var linkA1 = ws.Cell("A1").CreateHyperlink();
        linkA1.ExternalAddress = new Uri("http://example.com");

        var linkA2 = ws.Cell("A2").CreateHyperlink();
        linkA2.ExternalAddress = new Uri("http://google.com");

        // Original problem was that linkA1 was shifted to A2, but linkA2 wasn't yet shifted to A3.
        // Thus original dictionary threw "An item with the same key has already been added"
        Assert.DoesNotThrow(() => ws.Row(1).InsertRowsAbove(1));

        Assert.IsFalse(ws.Cell("A1").HasHyperlink);
        Assert.AreSame(linkA1, ws.Cell("A2").GetHyperlink());
        Assert.AreSame(linkA2, ws.Cell("A3").GetHyperlink());
    }

    [Test]
    public void Delete_link_removes_link_from_cell()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        var deleted = ws.Hyperlinks.Delete(link);

        // Assert
        Assert.IsTrue(deleted);
        Assert.AreEqual(ws.Style.Font.FontColor, ws.Cell("A1").Style.Font.FontColor);
        Assert.AreEqual(ws.Style.Font.Underline, ws.Cell("A1").Style.Font.Underline);
        Assert.IsNull(link.Container);
        Assert.IsFalse(ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out _));
    }

    [Test]
    public void Delete_link_for_address_deletes_the_link()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        var deleted = ws.Hyperlinks.Delete(ws.Cell("A1").Address);

        // Assert
        Assert.IsTrue(deleted);
        Assert.AreEqual(ws.Style.Font.FontColor, ws.Cell("A1").Style.Font.FontColor);
        Assert.AreEqual(ws.Style.Font.Underline, ws.Cell("A1").Style.Font.Underline);
        Assert.IsNull(link.Container);
        Assert.IsFalse(ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out _));
    }

    [Test]
    public void Delete_links_for_cell_address_without_link_doesnt_throw()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Act
        var wasDeleted = ws.Hyperlinks.Delete(ws.Cell("A1").Address);

        // Assert
        Assert.IsFalse(wasDeleted);
    }

    [Test]
    public void Delete_link_for_address_of_wrong_sheet_doesnt_delete_the_link()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var link = ws1.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        var ws2 = wb.AddWorksheet();

        // Act
        var deleted = ws1.Hyperlinks.Delete(ws2.Cell("A1").Address);

        // Assert
        Assert.IsFalse(deleted);
    }

    [Test]
    public void Get_returns_hyperlink_for_address()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        var foundLink = ws.Hyperlinks.Get(ws.Cell("A1").Address);

        // Assert
        Assert.AreSame(link, foundLink);
    }

    [Test]
    public void Get_throws_exception_when_address_doesnt_have_link()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        var otherSheet = wb.AddWorksheet();

        // Act + Assert
        foreach (var addressWithoutLink in new[] { ws.Cell("A2").Address, otherSheet.Cell("A1").Address })
        {
            Assert.Throws<KeyNotFoundException>(() => ws.Hyperlinks.Get(addressWithoutLink));
        }
    }

    [Test]
    public void Get_only_returns_links_from_correct_sheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var link = ws1.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        var ws2 = wb.AddWorksheet();
        var wrongSheetAddress = ws2.Cell("A1").Address;

        Assert.Throws<KeyNotFoundException>(() => ws1.Hyperlinks.Get(wrongSheetAddress));
    }

    [Test]
    public void TryGet_returns_hyperlink_for_address()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        var wasFound = ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out var foundLink);

        // Assert
        Assert.IsTrue(wasFound);
        Assert.AreSame(link, foundLink);
    }

    [Test]
    public void TryGet_doesnt_return_link_for_wrong_address()
    {
        // Arrange
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        var otherSheet = wb.AddWorksheet();

        // Act + Assert
        foreach (var addressWithoutLink in new[] { ws.Cell("A2").Address, otherSheet.Cell("A1").Address })
        {
            Assert.IsFalse(ws.Hyperlinks.TryGet(addressWithoutLink, out _));
        }
    }
}
