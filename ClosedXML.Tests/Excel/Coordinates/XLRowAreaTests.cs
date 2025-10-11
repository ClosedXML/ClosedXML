using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates;

[TestOf(typeof(XLRowArea))]
internal class XLRowAreaTests
{
    [TestCase(null)]
    [TestCase("")]
    public void Ctor_sheet_must_be_valid(string invalidSheetName)
    {
        Assert.That(
            () => new XLRowArea(invalidSheetName, 1),
            Throws.Exception.TypeOf<ArgumentException>());
    }

    [TestCase(-50)]
    [TestCase(0)]
    [TestCase(XLHelper.MaxRowNumber + 1)]
    [TestCase(int.MaxValue)]
    public void Ctor_row_number_must_be_valid(int invalidRowNumber)
    {
        Assert.That(
            () => new XLRowArea("some sheet", invalidRowNumber),
            Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
    }

    [TestCase("name", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 4, false)]
    [TestCase("some name", 1, "other name", 1, false)]
    public void Two_areas_are_compared_by_case_insensitive_sheet_name_and_row_number(string firstName, int firstRow, string secondName, int secondRow, bool areEqual)
    {
        var first = new XLRowArea(firstName, firstRow);
        var second = new XLRowArea(secondName, secondRow);
        Assert.That(first == second, Is.EqualTo(areEqual));
        Assert.That(first.GetHashCode() == second.GetHashCode(), Is.EqualTo(areEqual));
    }

    [Test]
    public void Area_property_returns_area_of_row()
    {
        var row = new XLRowArea("name", 4);
        var rowArea = row.Area;
        Assert.AreEqual(rowArea, new XLBookArea("name", new XLSheetRange(4, XLHelper.MinColumnNumber, 4, XLHelper.MaxColumnNumber)));
    }

    [TestCase("name", 4, "name!4:4")]
    [TestCase("some name", 4, "'some name'!4:4")]
    [TestCase("Joe's", 4, "'Joe''s'!4:4")]
    [TestCase("Joe", XLHelper.MaxRowNumber, "Joe!1048576:1048576")]
    public void ToString_returns_readable_reference(string name, int rowNumber, string expected)
    {
        var rowArea = new XLRowArea(name, rowNumber);
        Assert.AreEqual(expected, rowArea.ToString());
    }
}
