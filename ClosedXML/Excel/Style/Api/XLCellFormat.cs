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

    /// <summary>
    /// Formatting is updated for used cells within these areas. Unused cells are ignored.
    /// </summary>
    internal IReadOnlyList<XLBookArea> UsedAreas { get; init; } = Array.Empty<XLBookArea>();

    /// <summary>
    /// Formatting is updated for these columns. This doesn't update cells within the columns, only
    /// the columns themselves.
    /// </summary>
    internal IReadOnlyList<XLColumnArea> Columns { get; init; } = Array.Empty<XLColumnArea>();

    /// <summary>
    /// Formatting is updated for these rows. This doesn't update cells within the rows, only
    /// the rows themselves.
    /// </summary>
    internal IReadOnlyList<XLRowArea> Rows { get; init; } = Array.Empty<XLRowArea>();

    /// <summary>
    /// Formatting is updated for these worksheets. This doesn't update cells within the sheets, only
    /// the sheets and materialized rows and columns of the sheets.
    /// </summary>
    internal IReadOnlyList<string> Worksheets { get; init; } = Array.Empty<string>();

    internal XLFontCellFormat Font => new(this);

    internal static XLCellFormat ForColumn(XLColumn column)
    {
        var columnArea = column.Area;
        return new XLCellFormat(column.Worksheet.Workbook)
        {
            UsedAreas = new[] { columnArea.Area },
            Columns = new[] { columnArea }
        };
    }

    internal static XLCellFormat ForRow(XLRow row)
    {
        var rowArea = row.Area;
        return new XLCellFormat(row.Worksheet.Workbook)
        {
            UsedAreas = new[] { rowArea.Area },
            Rows = new[] { rowArea }
        };
    }

    internal static XLCellFormat ForWorksheet(XLWorksheet worksheet)
    {
        return new XLCellFormat(worksheet.Workbook)
        {
            UsedAreas = new[] { worksheet.Area },
            Worksheets = new[] { worksheet.Name }
        };
    }

    internal T Resolve<T>(Func<XLCellFormatValue, T?> selector)
        where T : struct
    {
        throw new NotImplementedException();
    }

    internal void ModifyFont<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        // TODO Styles: Apply to containers and deal with cross points
        var styles = _workbook.Styles;

        foreach (var sheetName in Worksheets)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            var originalFormat = worksheet.FormatValue ?? styles.DefaultFormat;
            var modifiedFormat = ModifyFormat(originalFormat);
            worksheet.FormatValue = modifiedFormat;

            var columns = worksheet.Internals.ColumnsCollection.Values;
            foreach (var column in columns)
                ApplyColRowFormat(column, worksheet);

            var rows = worksheet.Internals.RowsCollection.Values;
            foreach (var row in rows)
                ApplyColRowFormat(row, worksheet);
        }

        foreach (var columnArea in Columns)
        {
            if (!_workbook.TryGetWorksheet(columnArea.Name, out XLWorksheet worksheet))
                continue;

            var column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, worksheet);
        }

        foreach (var rowArea in Rows)
        {
            if (!_workbook.TryGetWorksheet(rowArea.Name, out XLWorksheet worksheet))
                continue;

            var row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, worksheet);
        }

        foreach (var (sheetName, area) in UsedAreas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(area, ModifyFormat, formatResolver.Resolve);
        }

        foreach (var (sheetName, area) in Areas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnAll(area, ModifyFormat, formatResolver.Resolve);
        }

        return;

        XLCellFormatValue ModifyFormat(XLCellFormatValue format)
        {
            var modifiedFont = styles.GetRegisteredFontFormat(format.Font, font => modifyFont(font, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Font = modifiedFont });
            return modifiedFormat;
        }

        void ApplyColRowFormat(IXLFormatContainer rowOrCol, XLWorksheet worksheet)
        {
            if (rowOrCol.FormatValue is not { } originalFormat)
                originalFormat = worksheet.FormatValue ?? styles.DefaultFormat;

            rowOrCol.FormatValue = ModifyFormat(originalFormat);
        }
    }
}
