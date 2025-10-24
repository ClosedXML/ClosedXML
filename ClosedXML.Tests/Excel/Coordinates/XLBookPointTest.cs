using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates;

internal class XLBookPointTest
{
    [TestCase(null)]
    [TestCase("")]
    public void Ctor_sheet_must_be_valid(string invalidSheetName)
    {
        Assert.That(
            () => new XLBookPoint(invalidSheetName, new XLSheetPoint(1, 1)),
            Throws.Exception.TypeOf<ArgumentException>());
    }

    [TestCase("sheet", 2, 5, "sheet", 2, 5, true)]
    [TestCase("SHEET", 2, 5, "sheet", 2, 5, true)]
    [TestCase("sheet", 2, 5, "sheet", 2, 6, false)]
    [TestCase("SHEET", 2, 5, "sheet", 3, 5, false)]
    [TestCase("some sheet", 2, 5, "other sheet", 2, 5, false)]
    public void Two_points_are_compared_by_case_insensitive_sheet_name_and_point_coordinates(string firstName, int firstRow, int firstColumn, string secondName, int secondRow, int secondColumn, bool areEqual)
    {
        var first = new XLBookPoint(firstName, firstRow, firstColumn);
        var second = new XLBookPoint(secondName, secondRow, secondColumn);
        Assert.That(first == second, Is.EqualTo(areEqual));
        Assert.That(first.GetHashCode() == second.GetHashCode(), Is.EqualTo(areEqual));
    }

    [TestCase("sheet", 1, 4, "sheet!D1")]
    [TestCase("Joe's", 47, 28, "'Joe''s'!AB47")]
    [TestCase("2025 Q1", XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber, "'2025 Q1'!XFD1048576")]
    public void ToString_returns_readable_reference(string name, int row, int column, string expected)
    {
        var bookPoint = new XLBookPoint(name, row, column);
        Assert.AreEqual(expected, bookPoint.ToString());
    }
}
