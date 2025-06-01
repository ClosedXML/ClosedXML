#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
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
    public void Cell_format_resolves_font_property_as_its_own_format_value_or_default_value(ResolveFormatTestCase cellFontTestCase)
    {
        // This test checks resolution of a font property value when a cell has a format.
        // * If cell has a format font property, that is resolved property.
        // * If cell has a format, but font property is not specified, it uses default value.
        //
        // Default value for font and other styles is kind of vague. So far it seems that
        // default value for missing font property is either a constant value that is essentially
        // a false (0)/black (color 000000)/first enum entry or it is taken from a normal
        // style if zero can't be a valid value (e.g. zero is not a valid font name or size).
        // Note that row or sheet formatting is never used. Only a constant value or a normal style.
        var styles = new XLWorkbookStyles();
        var actual = cellFontTestCase.InitActual();
        var rowOrColumn = cellFontTestCase.InitRowOrColumn();
        var sheet = cellFontTestCase.InitSheet();
        var normal = cellFontTestCase.InitNormal();
        var hierarchy = new FormatHierarchy(new[] { actual }, rowOrColumn, rowOrColumn, sheet, normal);
        var style = new XLCellFormat(hierarchy, styles);

        // Resolve font property
        var resolvedFontPropertyValue = cellFontTestCase.GetActual(style);

        // Assert that it's same as expected value
        Assert.AreEqual(cellFontTestCase.ExpectedValue, resolvedFontPropertyValue);
    }

    public static IEnumerable<object[]> GetFontDefaultValueTestCases()
    {
        return MakeFontFlagCases((b, f) => f with { Bold = (bool?)b }, x => x.Bold)
            .Concat(MakeFontFlagCases((i, f) => f with { Italic = (bool?)i }, x => x.Italic))
            .Concat(MakeFontNameCases())
            .Select(x => new object[] { x })
            .ToList();
    }

    private static IEnumerable<ResolveFormatTestCase> MakeFontFlagCases(Func<object?, XLFontFormatValue, XLFontFormatValue> setter, Func<XLFontCellFormat, object> getter)
    {
        // Font flags don't use normal style or any other fallback. If a cell font format doesn't have a value, it is just false.
        yield return MakeTestCase(setter, getter, true, null, null, null, true);
        yield return MakeTestCase(setter, getter, false, null, null, null, false);
        yield return MakeTestCase(setter, getter, null, null, null, null, false);
        yield return MakeTestCase(setter, getter, null, true, true, true, false);
    }

    private static IEnumerable<ResolveFormatTestCase> MakeFontNameCases()
    {
        // When cell has a font, use that font.
        yield return MakeTestCase((v, f) => f with { Name = (XLFontName?)v }, f => f.Name.Text, "Cell Font", "Row or Col Font", "Sheet Font", "Normal Font", (XLFontName)"Cell Font");

        // When cell doesn't have a font, use the font name form normal, even though there is a row and sheet in the hierarchy
        yield return MakeTestCase((v, f) => f with { Name = (XLFontName?)v }, f => f.Name.Text, null, "Row or Col Font", "Sheet Font", "Normal Font", (XLFontName)"Normal Font");
    }

    private static ResolveFormatTestCase MakeTestCase<T>(Func<object?, XLFontFormatValue, XLFontFormatValue> setter, Func<XLFontCellFormat, object> getter, T? cellValue, T? rowOrColValue, T? sheetValue, T? normalValue, T expected)
        where T : struct
    {
        return new ResolveFormatTestCase
        {
            ExpectedValue = expected,
            GetActual = x => getter(x.Font),
            Setter = setter,
            CellValue = cellValue,
            RowOrColValue = rowOrColValue,
            SheetValue = sheetValue,
            NormalValue = normalValue
        };
    }

    internal class ResolveFormatTestCase
    {
        public required object ExpectedValue { get; init; }

        public required Func<XLCellFormat, object> GetActual { get; init; }

        public required Func<object?, XLFontFormatValue, XLFontFormatValue> Setter { get; init; }

        public required object? CellValue { get; init; }

        public required object? RowOrColValue { get; init; }

        public required object? SheetValue { get; init; }

        public required object? NormalValue { get; init; }

        public IXLFormatContainer InitActual() => CreateFontContainer(CellValue);

        public IXLFormatContainer InitRowOrColumn() => CreateFontContainer(RowOrColValue);

        public IXLFormatContainer InitSheet() => CreateFontContainer(SheetValue);

        public IXLFormatContainer InitNormal() => CreateFontContainer(NormalValue);

        private FormatContainer CreateFontContainer(object? value)
        {
            var font = Setter(value, XLFontFormatValue.Empty);
            var format = XLCellFormatValue.Empty with { Font = font };
            return new FormatContainer { FormatValue = format };
        }
    }

    private class FormatContainer : IXLFormatContainer
    {
        public XLCellFormatValue? FormatValue { get; set; } = XLCellFormatValue.Empty;
    }
}
