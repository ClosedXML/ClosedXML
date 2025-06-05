#nullable enable
using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Excel.Formatting;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles;

/// <summary>
/// Tests of <see cref="XLCellFormat"/> API.
/// </summary>
[TestOf(typeof(XLCellFormat))]
[TestOf(typeof(IXLStyle))]
internal class XLCellFormatTests
{
    [TestCaseSource(nameof(GetFontDefaultValueTestCases))]
    public void Format_getters_resolve_cell_format_property_from_format_if_specified_or_from_default_format(XLCellFormat format, Func<XLCellFormat, object> getter, object expectedValue)
    {
        var formatPropertyValue = getter(format);

        Assert.AreEqual(expectedValue, formatPropertyValue);
    }

    [TestCaseSource(nameof(FormatHierarchySelectionTestCases))]
    public void Format_getters_use_closest_existing_format_in_hierarchy(XLCellFormat cellFormat, double expectedFontSize)
    {
        // There is a hierarchy, cell, row, column and normal style. When IXLStyle needs to
        // determine a format, it uses the first element in the hierarchy that has a format.
        // There is no format cascading.
        Assert.AreEqual(expectedFontSize, cellFormat.Font.Size);
    }

    public static IEnumerable<object[]> FormatHierarchySelectionTestCases()
    {
        // Make each element with a different font size so it's easy to check which format was selected.
        var cell18 = new FormatContainer { FormatValue = XLCellFormatValue.Empty with { Font = XLFontFormatValue.Empty with { Size = XLFontSize.FromPoints(18) } } };
        var row16 = new FormatContainer { FormatValue = XLCellFormatValue.Empty with { Font = XLFontFormatValue.Empty with { Size = XLFontSize.FromPoints(16) } } };
        var column14 = new FormatContainer { FormatValue = XLCellFormatValue.Empty with { Font = XLFontFormatValue.Empty with { Size = XLFontSize.FromPoints(14) } } };
        var normal12 = new FormatContainer { FormatValue = XLCellFormatValue.Empty with { Font = XLFontFormatValue.Empty with { Size = XLFontSize.FromPoints(12) } } };

        yield return MakeCellCase(cell18, null, null, normal12, 18);
        yield return MakeCellCase(null, null, null, normal12, 12);
        yield return MakeCellCase(null, row16, null, normal12, 16);
        yield return MakeCellCase(null, null, column14, normal12, 14);
        yield return MakeCellCase(null, row16, column14, normal12, 16);

        yield return MakeRowCase(row16, normal12, 16);
        yield return MakeRowCase(null, normal12, 12);

        yield return MakeColumnCase(row16, normal12, 16);
        yield return MakeColumnCase(null, normal12, 12);
        yield break;

        static object[] MakeCellCase(FormatContainer? cell, FormatContainer? row, FormatContainer? column, FormatContainer normal, int expectedFontSize)
        {
            cell ??= new FormatContainer { FormatValue = null };
            return new object[] { new XLCellFormat(FormatHierarchy.ForCell(cell, row, column, normal), new XLWorkbookStyles()), expectedFontSize };
        }

        static object[] MakeRowCase(FormatContainer? row, FormatContainer normal, int expectedFontSize)
        {
            row ??= new FormatContainer { FormatValue = null };
            return new object[] { new XLCellFormat(FormatHierarchy.ForRow(row, normal), new XLWorkbookStyles()), expectedFontSize };
        }

        static object[] MakeColumnCase(FormatContainer? column, FormatContainer normal, int expectedFontSize)
        {
            column ??= new FormatContainer { FormatValue = null };
            return new object[] { new XLCellFormat(FormatHierarchy.ForColumn(column, normal), new XLWorkbookStyles()), expectedFontSize };
        }
    }

    public static IEnumerable<object[]> GetFontDefaultValueTestCases()
    {
        yield return MakeCase(f => f with { Name = "Arial" }, f => f.Font.Name, "Arial");
        yield return MakeCase(f => f with { Name = null }, f => f.Font.Name, "Default Font");
        yield return MakeCase(f => f with { Size = XLFontSize.FromPoints(7) }, f => f.Font.Size, 7);
        yield return MakeCase(f => f with { Size = null }, f => f.Font.Size, 15);
        yield return MakeCase(f => f with { Bold = true }, f => f.Font.Bold, true);
        yield return MakeCase(f => f with { Bold = null }, f => f.Font.Bold, false);

    }

    private static object[] MakeCase(Func<XLFontFormatValue, XLFontFormatValue> modify, Func<XLCellFormat, object> getter, object expectedValue)
    {
        var styles = new XLWorkbookStyles();
        var cell = new FormatContainer
        {
            FormatValue = XLCellFormatValue.Empty with { Font = modify(XLFontFormatValue.Empty) }
        };
        var normal = new FormatContainer
        {
            FormatValue = styles.DefaultFormat
        };
        styles.DefaultFormat = styles.DefaultFormat with { Font = styles.DefaultFormat.Font! with { Name = "Default Font", Size = XLFontSize.FromPoints(15) } };
        var format = new XLCellFormat(FormatHierarchy.ForCell(cell, null, null, normal), styles);
        return new[] { format, getter, expectedValue };
    }

    private class FormatContainer : IXLFormatContainer
    {
        public required XLCellFormatValue? FormatValue { get; set; }
    }
}
