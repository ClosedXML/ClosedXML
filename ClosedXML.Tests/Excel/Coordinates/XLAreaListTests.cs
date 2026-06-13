using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates;

[TestFixture]
internal class XLAreaListTests
{
    [TestCase("A1:C3", "A1", "B1:C3 A2:A4")]
    [TestCase("A1:C3", "B1", "A1:A3 C1:C3 B2:B4")]
    [TestCase("A1:C3", "C1", "A1:B3 C2:C4")]
    [TestCase("A1:C3", "A2", "A1:C1 B2:C3 A2:A4")]
    [TestCase("A1:C3", "B2", "A1:C1 A2:A3 C2:C3 B2:B4")]
    [TestCase("A1:C3", "C2", "A1:C1 A2:B3 C2:C4")]
    [TestCase("A1:C3", "A3", "A1:C2 B3:C3 A3:A4")]
    [TestCase("A1:C3", "B3", "A1:C2 A3 C3 B3:B4")]
    [TestCase("A1:C3", "C3", "A1:C2 A3:B3 C3:C4")]

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
        var list = new XLAreaList(Area.Parse(areaList));
        var result = list.InsertAndShiftDown(Area.Parse(insertedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [Test]
    public void InsertAndShiftDown_baseline_comparison()
    {
        // Compare the result of the method with the behavior of CFs Applied To field collected from Excel
        foreach (var (original, insertArea, expectedResult) in GetBaselineData("Other.ConditionalFormats.insert-and-shift-down-cf-baseline.txt"))
        {
            var result = original.InsertAndShiftDown(insertArea);
            Assert.AreEqual(expectedResult.ToSpaceList(), result.ToSpaceList());
        }
    }

    [TestCase("A1:C3", "A1", "A2:C3 B1:D1")]
    [TestCase("A1:C3", "B1", "A2:C3 A1 C1:D1")]
    [TestCase("A1:C3", "C1", "A2:C3 A1:B1 D1")]
    [TestCase("A1:C3", "A2", "A1:C1 A3:C3 B2:D2")]
    [TestCase("A1:C3", "B2", "A1:C1 A3:C3 A2 C2:D2")]
    [TestCase("A1:C3", "C2", "A1:C1 A3:C3 A2:B2 D2")]
    [TestCase("A1:C3", "A3", "A1:C2 B3:D3")]
    [TestCase("A1:C3", "B3", "A1:C2 A3 C3:D3")]
    [TestCase("A1:C3", "C3", "A1:C2 A3:B3 D3")]

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
        var list = new XLAreaList(new List<Area> { Area.Parse(areaList) });
        var result = list.InsertAndShiftRight(Area.Parse(insertedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [TestCase("A1:C3", "A1", "B1:C3 A1:A2")]
    [TestCase("A1:C3", "B1", "A1:A3 C1:C3 B1:B2")]
    [TestCase("A1:C3", "C1", "A1:B3 C1:C2")]
    [TestCase("A1:C3", "A2", "A1:C1 B2:C3 A2")]
    [TestCase("A1:C3", "B2", "A1:C1 A2:A3 C2:C3 B2")]
    [TestCase("A1:C3", "C2", "A1:C1 A2:B3 C2")]
    [TestCase("A1:C3", "A3", "A1:C2 B3:C3")]
    [TestCase("A1:C3", "B3", "A1:C2 A3 C3")]
    [TestCase("A1:C3", "C3", "A1:C2 A3:B3")]

    [TestCase("B1:D3", "A1:A3", "B1:D3")] // Delete on the left side - don't move
    [TestCase("A2:C4", "A1:C1", "A1:C3")] // Delete on top side - shift
    [TestCase("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [TestCase("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move

    [TestCase("A1:A3", "A1:D5", "")] // Delete completely
    [TestCase("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public void DeleteAndShiftUp(string areaList, string deletedArea, string expected)
    {
        var list = new XLAreaList(new List<Area> { Area.Parse(areaList) });
        var result = list.DeleteAndShiftUp(Area.Parse(deletedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [TestCase("A1:C3", "A1", "A2:C3 A1:B1")]
    [TestCase("A1:C3", "B1", "A2:C3 A1 B1")]
    [TestCase("A1:C3", "C1", "A2:C3 A1:B1")]
    [TestCase("A1:C3", "A2", "A1:C1 A3:C3 A2:B2")]
    [TestCase("A1:C3", "B2", "A1:C1 A3:C3 A2 B2")]
    [TestCase("A1:C3", "C2", "A1:C1 A3:C3 A2:B2")]
    [TestCase("A1:C3", "A3", "A1:C2 A3:B3")]
    [TestCase("A1:C3", "B3", "A1:C2 A3 B3")]
    [TestCase("A1:C3", "C3", "A1:C2 A3:B3")]

    [TestCase("B1:D3", "A1:A3", "A1:C3")] // Delete on the left side - shift
    [TestCase("A2:C4", "A1", "A2:C4")] // Delete on top side - don't move
    [TestCase("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [TestCase("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move

    [TestCase("A1:A3", "A1:D5", "")] // Delete completely
    [TestCase("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public void DeleteAndShiftLeft(string areaList, string deletedArea, string expected)
    {
        var list = new XLAreaList(new List<Area> { Area.Parse(areaList) });
        var result = list.DeleteAndShiftLeft(Area.Parse(deletedArea));

        Assert.AreEqual(expected, result.ToSpaceList());
    }

    [TestCase("A1", "A1", ExpectedResult = true)]
    [TestCase("A1:C3", "B2", ExpectedResult = true)]
    [TestCase("B2:C3", "A2", ExpectedResult = false)]
    [TestCase("A1:C2 B3:C3", "A3", ExpectedResult = false)]
    public bool IntersectsWith_determines_intersection_with_any_area(string areaListText, string areaText)
    {
        var areaList = Parse(areaListText);
        var area = Area.Parse(areaText);
        return areaList.IntersectsWith(area);
    }

    [TestCase("A1", "A1", ExpectedResult = "A1")]
    [TestCase("A1:C3", "B2", ExpectedResult = "A1:C3")]
    [TestCase("A1:C3", "B2:D4", ExpectedResult = "A1:C3")]
    [TestCase("A1 C1", "A1:C1", ExpectedResult = "A1 C1")]
    [TestCase("A1 C1", "B1", ExpectedResult = "")]
    [TestCase("A1 C1", "B1:D2", ExpectedResult = "C1")]
    public string IntersectingWith_returns_areas_intersecting_with_the_other_area(string areaListText, string areaText)
    {
        var areaList = Parse(areaListText);
        var area = Area.Parse(areaText);
        return areaList.IntersectingWith(area).ToSpaceList();
    }

    [TestCase("A1", "B1", ExpectedResult = "A1")]
    [TestCase("A1:E5", "C3:C4", ExpectedResult = "A1:E2 A5:E5 A3:B4 D3:E4")]
    [TestCase("B2:C5 B9 C4:D7", "C4:C5", ExpectedResult = "B2:C3 B4:B5 B9 C6:D7 D4:D5")]
    public string Excluding_returns_area_list_without_excluded(string areaListText, string excludedAreaText)
    {
        var areaList = Parse(areaListText);
        var excludedArea = Area.Parse(excludedAreaText);
        return areaList.Excluding(excludedArea).ToSpaceList();
    }

    [TestCase("A1", "A1", "A1", ExpectedResult = "A1")] // Copy from same point to the same point
    [TestCase("A1", "B5", "A1", ExpectedResult = "B5")] // Copy to different point
    [TestCase("B2", "D2", "A1:C3", ExpectedResult = "E3")] // The intersected area is not in corner and shifted doesn't start at the target point
    [TestCase("D3:G6", "A1", "E4:F5", ExpectedResult = "A1:B2")]
    [TestCase("B2", XLHelper.LastSheetAddress, "A1:C3", ExpectedResult = null)] // Copied area is out of sheet. Rare, but can happen.
    public string TryCopyAreaTo_return_list_of_intersecting_areas_shifted_to_target(string areaListText, string targetPointText, string areaToCopyText)
    {
        var areaList = Parse(areaListText);
        var targetPoint = Point.Parse(targetPointText);
        var areaToCopy = Area.Parse(areaToCopyText);
        return areaList.TryCopyAreaTo(targetPoint, areaToCopy, out var result) ? result.ToSpaceList() : null;
    }

    private static XLAreaList Parse(string spaceList)
    {
        var list = new List<Area>();
        foreach (var reference in spaceList.Split(' '))
            list.Add(Area.Parse(reference));

        return new XLAreaList(list);
    }

    private static IEnumerable<(XLAreaList, Area, XLAreaList)> GetBaselineData(string resourcePath)
    {
        using var stream = TestHelper.GetStreamFromResource(resourcePath);
        using var streamReader = new StreamReader(stream);
        while (streamReader.ReadLine() is { } line)
        {
            var fields = line.Split(',');
            var original = Parse(fields[0]);
            var area = Area.Parse(fields[1]);
            var expectedResult = Parse(fields[2]);
            yield return (original, area, expectedResult);
        }
    }
}
