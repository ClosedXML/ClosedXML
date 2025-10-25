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

    private XLCellFormat(XLWorkbook workbook)
    {
        _workbook = workbook;
    }

    internal XLFontCellFormat Font => new XLFontCellFormat(this);

    /// <summary>
    /// Cell areas in a workbook that should be updated when format is changed, e.g. when we have
    /// a format API object for a row container, the area are all cells of the row. It must be
    /// an area, so we can satisfy the <see cref="IXLBorder.OutsideBorder"/> and
    /// <see cref="IXLBorder.InsideBorder"/> property setters.
    /// </summary>
    private IReadOnlyList<XLBookArea> Areas { get; init; } = Array.Empty<XLBookArea>();

    /// <summary>
    /// Formatting is updated for used cells within these areas. Unused cells are ignored.
    /// </summary>
    private IReadOnlyList<XLBookArea> UsedAreas { get; init; } = Array.Empty<XLBookArea>();

    /// <summary>
    /// Formatting is updated for these columns. This doesn't update cells within the columns, only
    /// the columns themselves.
    /// </summary>
    private IReadOnlyList<XLColumnArea> Columns { get; init; } = Array.Empty<XLColumnArea>();

    /// <summary>
    /// Formatting is updated for these rows. This doesn't update cells within the rows, only
    /// the rows themselves.
    /// </summary>
    private IReadOnlyList<XLRowArea> Rows { get; init; } = Array.Empty<XLRowArea>();

    /// <summary>
    /// Formatting is updated for these worksheets. This doesn't update cells within the sheets, only
    /// the sheets and materialized rows and columns of the sheets.
    /// </summary>
    private IReadOnlyList<string> Worksheets { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Should the formatting be updated for the default format of a workbook (plus cascade to all
    /// formats below, containers and areas).
    /// </summary>
    private bool DefaultFormat { get; init; }

    internal static XLCellFormat ForCell(XLCell cell)
    {
        return new XLCellFormat(cell.Worksheet.Workbook)
        {
            Areas = new[] { new XLBookArea(cell.Worksheet.Name, new XLSheetRange(cell.SheetPoint)) }
        };
    }

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

    internal static XLCellFormat ForWorkbook(XLWorkbook workbook)
    {
        return new XLCellFormat(workbook)
        {
            DefaultFormat = true
        };
    }

    internal T Resolve<T>(Func<XLCellFormatValue, T?> selector)
        where T : struct
    {
        throw new NotImplementedException();
    }

    internal void ModifyFont<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        // TODO Styles: Deal with cross points
        var styles = _workbook.Styles;
        var modifyFormat = (XLCellFormatValue format) =>
        {
            var modifiedFont = styles.GetRegisteredFontFormat(format.Font, font => modifyFont(font, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Font = modifiedFont });
            return modifiedFormat;
        };

        if (DefaultFormat)
        {
            styles.DefaultFormat = modifyFormat(styles.DefaultFormat);
            foreach (var worksheet in _workbook.WorksheetsInternal)
            {
                ApplyToWorksheet(worksheet, modifyFormat, styles);
            }
        }

        foreach (var sheetName in Worksheets)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            ApplyToWorksheet(worksheet, modifyFormat, styles);
        }

        foreach (var columnArea in Columns)
        {
            if (!_workbook.TryGetWorksheet(columnArea.Name, out XLWorksheet worksheet))
                continue;

            var column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, modifyFormat, worksheet);
        }

        foreach (var rowArea in Rows)
        {
            if (!_workbook.TryGetWorksheet(rowArea.Name, out XLWorksheet worksheet))
                continue;

            var row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, modifyFormat, worksheet);
        }

        foreach (var (sheetName, area) in UsedAreas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            ApplyToUsed(area, modifyFormat, worksheet);
        }

        foreach (var (sheetName, area) in Areas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            ApplyToAll(area, modifyFormat, worksheet);
        }
    }

    private static void ApplyToWorksheet(XLWorksheet worksheet, Func<XLCellFormatValue, XLCellFormatValue> modifyFormat, XLWorkbookStyles styles)
    {
        var originalFormat = worksheet.FormatValue ?? styles.DefaultFormat;
        var modifiedFormat = modifyFormat(originalFormat);
        worksheet.FormatValue = modifiedFormat;

        var columns = worksheet.Internals.ColumnsCollection.Values;
        foreach (var column in columns)
            ApplyColRowFormat(column, modifyFormat, worksheet);

        var rows = worksheet.Internals.RowsCollection.Values;
        foreach (var row in rows)
            ApplyColRowFormat(row, modifyFormat, worksheet);

        ApplyToUsed(XLSheetRange.Full, modifyFormat, worksheet);
    }

    private static void ApplyColRowFormat(IXLFormatContainer rowOrCol, Func<XLCellFormatValue, XLCellFormatValue> modifyFormat, XLWorksheet worksheet)
    {
        if (rowOrCol.FormatValue is not { } originalFormat)
            originalFormat = worksheet.FormatValue ?? worksheet.Workbook.Styles.DefaultFormat;

        rowOrCol.FormatValue = modifyFormat(originalFormat);
    }

    private static void ApplyToUsed(XLSheetRange area, Func<XLCellFormatValue, XLCellFormatValue> modifyFormat, XLWorksheet worksheet)
    {
        var formatResolver = new FormatResolver(worksheet);
        var cellsCollection = worksheet.Internals.CellsCollection;
        cellsCollection.ApplyFormatOnUsed(area, modifyFormat, formatResolver.Resolve);
    }

    private static void ApplyToAll(XLSheetRange area, Func<XLCellFormatValue, XLCellFormatValue> modifyFormat, XLWorksheet worksheet)
    {
        var formatResolver = new FormatResolver(worksheet);
        var cellsCollection = worksheet.Internals.CellsCollection;
        cellsCollection.ApplyFormatOnAll(area, modifyFormat, formatResolver.Resolve);
    }
}
