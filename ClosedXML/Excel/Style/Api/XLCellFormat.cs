using System;
using System.Collections.Generic;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// Unlike the <see cref="XLStyle"/>, the <see cref="XLCellFormat"/> one modifies formatting
/// in a <see cref="XLWorkbookStyles"/>.
/// </summary>
internal class XLCellFormat
{
    private readonly XLWorkbook _workbook;

    public XLCellFormat(XLWorkbook workbook)
    {
        _workbook = workbook;
    }

    /// <summary>
    /// Cell areas in a workbook that should be updated when format is changed, e.g. when we have
    /// a format API object for a row container, the area are all cells of the row. It must be
    /// an area, so we can satisfy the <see cref="IXLBorder.OutsideBorder"/> and
    /// <see cref="IXLBorder.InsideBorder"/> property setters.
    /// </summary>
    internal IReadOnlyList<XLBookArea> Areas { get; init; } = Array.Empty<XLBookArea>();

    internal XLFontCellFormat Font => new(this);

    internal T Resolve<T>(Func<XLCellFormatValue, T?> selector)
        where T : struct
    {
        throw new NotImplementedException();
    }

    internal void ModifyFont<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        var styles = _workbook.Styles;
        foreach (var (sheetName, area) in Areas)
        {
            // Worksheet could have been deleted -> skip
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var formatSlice = worksheet.Internals.CellsCollection.FormatSlice;
            formatSlice.ApplyDeterministic(area, format =>
            {
                var modifiedFont = styles.GetRegisteredFontFormat(format.Font, font => modifyFont(font, value));
                var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Font = modifiedFont });
                return modifiedFormat;
            }, formatResolver.Resolve);
        }

        // TODO: Apply to used areas, containers and deal with cross points
    }
}
