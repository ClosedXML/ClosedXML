using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates;

[TestFixture]
internal class XLAreaListTests
{
    [TestCase("A1:C3", "A1", "B1:C3 A2:A4")]
    [TestCase("A1:C3", "B1", "A1:A3 C1:C3 B2:B4")]
    [TestCase("A1:C3", "C1", "A1:B3 C2:C4")]
    [TestCase("A1:C3", "A2", "A1 B1:C3 A3:A4")]
    [TestCase("A1:C3", "B2", "A1:A3 B1 C1:C3 B3:B4")]
    [TestCase("A1:C3", "C2", "A1:B3 C1 C3:C4")]
    [TestCase("A1:C3", "A3", "A1:A2 B1:C3 A4")]
    [TestCase("A1:C3", "B3", "A1:A3 B1:B2 C1:C3 B4")]
    [TestCase("A1:C3", "C3", "A1:B3 C1:C2 C4")]

    [TestCase("B1:D3", "A1:A3", "B1:D3")] // Insert to left side - don't move
    [TestCase("A2:C4", "A1:C1", "A3:C5")] // Insert to top side - shift
    [TestCase("A2:C4", "A2:C2", "A3:C5")] // Insert to top edge - shift
    [TestCase("A2:C4", "A1", "B2:C4 A3:A5")] // Insert to top side - shift
    [TestCase("A1:C3", "D1:D3", "A1:C3")] // Insert to right side - don't move
    [TestCase("A1:C3", "A4:C5", "A1:C5")] // Insert to bottom edge - extend
    [TestCase("A1:C3", "A4", "A1:C3 A4")] // Insert to bottom side - extend
    [TestCase("A1:C3", "B4:E5", "A1:C3 B4:C5")] // Insert to bottom edge (inserted area is out of bounds of the area) - extend

    [TestCase("A1048576", "A1048576", "")] // Push out of sheet
    [TestCase("A1048575:A1048576", "A1048575", "A1048576")] // Partially push out of sheet
    [TestCase("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public void InsertAndShiftDown(string areaList, string insertedArea, string expected)
    {
        var list = new XLAreaList(new List<XLSheetRange> { XLSheetRange.Parse(areaList) });
        var result = list.InsertAndShiftDown(XLSheetRange.Parse(insertedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [TestCase("A1:C3", "A1", "A2:C3 B1:D1")]
    [TestCase("A1:C3", "B1", "A1:A3 B2:C3 C1:D1")]
    [TestCase("A1:C3", "C1", "A1:B3 C2:C3 D1")]
    [TestCase("A1:C3", "A2", "A1:C1 A3:C3 B2:D2")]
    [TestCase("A1:C3", "B2", "A1:A3 B1:C1 B3:C3 C2:D2")]
    [TestCase("A1:C3", "C2", "A1:B3 C1 C3 D2")]
    [TestCase("A1:C3", "A3", "A1:C2 B3:D3")]
    [TestCase("A1:C3", "B3", "A1:A3 B1:C2 C3:D3")]
    [TestCase("A1:C3", "C3", "A1:B3 C1:C2 D3")]

    [TestCase("A1:C3", "A1:A3", "B1:D3")] // Insert to left edge - shift, don't extend
    [TestCase("A2:C4", "A1", "A2:C4")] // Insert to top side - don't move
    [TestCase("A1:C3", "D1:D3", "A1:D3")] // Insert to right edge - extend
    [TestCase("A1:C3", "D2:E10", "A1:C3 D2:E3")] // Insert to right edge (inserted area is out of bounds of the area) - extend
    [TestCase("A1:C3", "E1:E3", "A1:C3")] // Insert to right side  - don't move
    [TestCase("A1:C3", "A4", "A1:C3")] // Insert to bottom side  - don't move

    [TestCase("XFD1", "XFD1", "")] // Push out of sheet
    [TestCase("XFC1:XFD1", "XFC1", "XFD1")] // Partially push out of sheet
    [TestCase("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public void InsertAndShiftRight(string areaList, string insertedArea, string expected)
    {
        var list = new XLAreaList(new List<XLSheetRange> { XLSheetRange.Parse(areaList) });
        var result = list.InsertAndShiftRight(XLSheetRange.Parse(insertedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [TestCase("A1:C3", "A1", "B1:C3 A1:A2")]
    [TestCase("A1:C3", "B1", "A1:A3 C1:C3 B1:B2")]
    [TestCase("A1:C3", "C1", "A1:B3 C1:C2")]
    [TestCase("A1:C3", "A2", "A1 B1:C3 A2")]
    [TestCase("A1:C3", "B2", "A1:A3 B1 C1:C3 B2")]
    [TestCase("A1:C3", "C2", "A1:B3 C1 C2")]
    [TestCase("A1:C3", "A3", "A1:A2 B1:C3")]
    [TestCase("A1:C3", "B3", "A1:A3 B1:B2 C1:C3")]
    [TestCase("A1:C3", "C3", "A1:B3 C1:C2")]

    [TestCase("B1:D3", "A1:A3", "B1:D3")] // Delete on the left side - don't move
    [TestCase("A2:C4", "A1:C1", "A1:C3")] // Delete on top side - shift
    [TestCase("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [TestCase("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move

    [TestCase("A1:A3", "A1:D5", "")] // Delete completely
    [TestCase("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public void DeleteAndShiftUp(string areaList, string deletedArea, string expected)
    {
        var list = new XLAreaList(new List<XLSheetRange> { XLSheetRange.Parse(areaList) });
        var result = list.DeleteAndShiftUp(XLSheetRange.Parse(deletedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [TestCase("A1:C3", "A1", "A2:C3 A1:B1")]
    [TestCase("A1:C3", "B1", "A1:A3 B2:C3 B1")]
    [TestCase("A1:C3", "C1", "A1:B3 C2:C3")]
    [TestCase("A1:C3", "A2", "A1:C1 A3:C3 A2:B2")]
    [TestCase("A1:C3", "B2", "A1:A3 B1:C1 B3:C3 B2")]
    [TestCase("A1:C3", "C2", "A1:B3 C1 C3")]
    [TestCase("A1:C3", "A3", "A1:C2 A3:B3")]
    [TestCase("A1:C3", "B3", "A1:A3 B1:C2 B3")]
    [TestCase("A1:C3", "C3", "A1:B3 C1:C2")]

    [TestCase("B1:D3", "A1:A3", "A1:C3")] // Delete on the left side - shift
    [TestCase("A2:C4", "A1", "A2:C4")] // Delete on top side - don't move
    [TestCase("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [TestCase("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move

    [TestCase("A1:A3", "A1:D5", "")] // Delete completely
    [TestCase("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public void DeleteAndShiftLeft(string areaList, string deletedArea, string expected)
    {
        var list = new XLAreaList(new List<XLSheetRange> { XLSheetRange.Parse(areaList) });
        var result = list.DeleteAndShiftLeft(XLSheetRange.Parse(deletedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }
}
