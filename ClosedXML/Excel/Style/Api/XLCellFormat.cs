using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// Unlike the <see cref="XLStyle"/>, the <see cref="XLCellFormat"/> one modifies formatting
/// in a <see cref="XLWorkbookStyles"/>.
/// </summary>
internal partial class XLCellFormat
{
    private readonly XLWorkbook _workbook;
    private readonly Hierarchy _formatValue;

    private XLCellFormat(XLWorkbook workbook, Hierarchy formatValue)
    {
        _workbook = workbook;
        _formatValue = formatValue;
    }

    internal XLNumberCellFormat NumberFormat => new(this);

    internal XLFontCellFormat Font => new(this);

    internal XLFillCellFormat Fill => new(this);
    
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
        var workbook = cell.Worksheet.Workbook;
        var sheetName = cell.Worksheet.Name;
        var cellPoint = cell.SheetPoint;
        var formatValue = new Hierarchy(workbook, sheetName, cellPoint.Column, cellPoint.Row, cellPoint);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = new[] { new XLBookArea(sheetName, new XLSheetRange(cellPoint)) }
        };
    }

    internal static XLCellFormat ForColumn(XLColumn column)
    {
        var workbook = column.Worksheet.Workbook;
        var columnArea = column.Area;
        var formatValue = new Hierarchy(workbook, columnArea.Name, columnArea.ColumNumber, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { columnArea.Area },
            Columns = new[] { columnArea }
        };
    }

    internal static XLCellFormat ForColumns(XLWorkbook workbook, XLWorksheet? formatValueSheet, IReadOnlyList<XLColumnArea> columns)
    {
        var formatValue = new Hierarchy(workbook, formatValueSheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Columns = columns,
            UsedAreas = columns.Select(x => x.Area).ToArray()
        };
    }

    internal static XLCellFormat ForRow(XLRow row)
    {
        var workbook = row.Worksheet.Workbook;
        var rowArea = row.Area;
        var formatValue = new Hierarchy(workbook, rowArea.Name, null, rowArea.RowNumber, null);
        return new XLCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { rowArea.Area },
            Rows = new[] { rowArea }
        };
    }

    internal static XLCellFormat ForRows(XLWorkbook workbook, XLWorksheet? formatValueSheet, IReadOnlyList<XLRowArea> rows)
    {
        var formatValue = new Hierarchy(workbook, formatValueSheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Rows = rows,
            UsedAreas = rows.Select(x => x.Area).ToArray()
        };
    }

    internal static XLCellFormat ForWorksheet(XLWorksheet worksheet)
    {
        var workbook = worksheet.Workbook;
        var formatValue = new Hierarchy(workbook, worksheet.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { worksheet.Area },
            Worksheets = new[] { worksheet.Name }
        };
    }

    internal static XLCellFormat ForWorkbook(XLWorkbook workbook)
    {
        var formatValue = new Hierarchy(workbook, null, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            DefaultFormat = true
        };
    }

    internal static XLCellFormat ForCells(XLWorkbook workbook, IReadOnlyList<XLBookArea> areas, XLWorksheet? sheet)
    {
        var formatValue = new Hierarchy(workbook, sheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = areas
        };
    }

    internal static XLCellFormat ForRange(XLWorksheet sheet, XLRangeAddress rangeAddress)
    {
        var workbook = sheet.Workbook;
        var formatValue = new Hierarchy(workbook, sheet.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = new[] { XLBookArea.From(rangeAddress) }
        };
    }

    internal static XLCellFormat ForTableRows(XLWorksheet sheet, XLBookArea[] rowAreas)
    {
        var workbook = sheet.Workbook;
        var formatValue = new Hierarchy(workbook, sheet.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = rowAreas
        };
    }

    internal T Resolve<T>(Func<XLCellFormatValue, T> selector)
    {
        var format = _formatValue.Resolve();
        return selector(format);
    }

    internal void ModifyNumberFormat(string numberFormat)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedNumberFormat = styles.GetRegisteredNumberFormat(numberFormat);
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { NumberFormat = modifiedNumberFormat });
            return modifiedFormat;
        });
    }

    internal void ModifyFont<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedFont = styles.GetRegisteredFontFormat(format.Font, font => modifyFont(font, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Font = modifiedFont });
            return modifiedFormat;
        });
    }

    internal void ModifyFill<TProperty>(Func<XLFillFormatValue, TProperty, XLFillFormatValue> modifyFill, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedFill = styles.GetRegisteredFillFormat(format.Fill, fill => modifyFill(fill, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Fill = modifiedFill });
            return modifiedFormat;
        });
    }

    private void Modify(Func<XLCellFormatValue, XLCellFormatValue> modifyFormat)
    {
        // TODO Styles: Deal with cross points
        var styles = _workbook.Styles;
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

    /// <summary>
    /// A format value resolution hierarchy for a range API object. Each range API type needs
    /// to set proper fallbacks through ctor.
    /// </summary>
    private readonly record struct Hierarchy
    {
        private readonly XLWorkbook _workbook;
        private readonly string? _sheetName;
        private readonly int? _columnNumber;
        private readonly int? _rowNumber;
        private readonly XLSheetPoint? _point;

        public Hierarchy(XLWorkbook workbook, string? sheetName, int? columnNumber, int? rowNumber, XLSheetPoint? point)
        {
            _workbook = workbook;
            _sheetName = sheetName;
            _columnNumber = columnNumber;
            _rowNumber = rowNumber;
            _point = point;
        }

        private XLCellFormatValue DefaultFormat => _workbook.Styles.DefaultCellFormat;

        internal XLCellFormatValue Resolve()
        {
            var isForWorkbook = _sheetName is null;
            if (isForWorkbook)
                return DefaultFormat;

            // First, make sure the sheet exists
            if (!_workbook.TryGetWorksheet(_sheetName, out XLWorksheet sheet))
                return DefaultFormat;

            if (_point is { } point)
            {
                var formatSlice = sheet.Internals.CellsCollection.FormatSlice;
                var cellFormat = formatSlice.GetFormat(point);
                if (cellFormat is not null)
                    return cellFormat;
            }

            if (_rowNumber is { } rowNumber)
            {
                var rowsCollection = sheet.Internals.RowsCollection;
                if (rowsCollection.TryGetValue(rowNumber, out var row) &&
                    row.FormatValue is { } rowFormat)
                {
                    return rowFormat;
                }
            }

            if (_columnNumber is { } columnNumber)
            {
                var columnsCollection = sheet.Internals.ColumnsCollection;
                if (columnsCollection.TryGetValue(columnNumber, out var column) &&
                    column.FormatValue is { } columnFormat)
                {
                    return columnFormat;
                }
            }

            if (sheet.FormatValue is { } sheetFormat)
                return sheetFormat;

            return DefaultFormat;
        }
    }
}
