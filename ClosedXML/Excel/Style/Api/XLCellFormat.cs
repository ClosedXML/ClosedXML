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

    internal XLBorderCellFormat Border => new(this);

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
    /// the columns themselves. The values are unique columns, sorted by column number in ascending
    /// order.
    /// </summary>
    private XLColumnArea[] Columns { get; init; } = Array.Empty<XLColumnArea>();

    /// <summary>
    /// Formatting is updated for these rows. This doesn't update cells within the rows, only
    /// the rows themselves. The values are unique rows, sorted by row number in ascending order.
    /// </summary>
    private XLRowArea[] Rows { get; init; } = Array.Empty<XLRowArea>();

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

    /// <summary>
    /// A flag indicating API object is for XLCells. Unlike other range API objects, the XLCells
    /// has a non-standard outside/inside borders behavior.
    /// </summary>
    private bool IsCells { get; init; }

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

    internal static XLCellFormat ForColumns(XLWorkbook workbook, XLWorksheet? formatValueSheet, IEnumerable<XLColumn> columns)
    {
        var columnAreas = columns.Select(x => x.Area).Distinct().OrderBy(x => x.ColumNumber).ToArray();
        var formatValue = new Hierarchy(workbook, formatValueSheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Columns = columnAreas,
            UsedAreas = columnAreas.Select(x => x.Area).ToArray()
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

    internal static XLCellFormat ForRows(XLWorkbook workbook, XLWorksheet? formatValueSheet, IEnumerable<XLRow> rows)
    {
        var rowAreas = rows.Select(x => x.Area).Distinct().OrderBy(x => x.RowNumber).ToArray();
        var formatValue = new Hierarchy(workbook, formatValueSheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Rows = rowAreas,
            UsedAreas = rowAreas.Select(x => x.Area).ToArray()
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

    internal static XLCellFormat ForAreas(XLWorkbook workbook, IReadOnlyList<XLBookArea> areas, XLWorksheet? sheet)
    {
        var formatValue = new Hierarchy(workbook, sheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = areas
        };
    }

    internal static XLCellFormat ForCells(XLWorkbook workbook, IReadOnlyList<XLBookArea> areas, XLWorksheet? sheet)
    {
        var formatValue = new Hierarchy(workbook, sheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = areas,
            IsCells = true
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

    internal void ModifyBorder<TProperty>(Func<XLBorderFormatValue, TProperty, XLBorderFormatValue> modifyBorder, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedBorder = styles.GetRegisteredBorderFormat(format.Border, border => modifyBorder(border, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Border = modifiedBorder });
            return modifiedFormat;
        });
    }

    internal void ModifyOuterBorder<TProperty>(Func<XLBorderLine, TProperty, XLBorderLine> modify, TProperty value)
    {
        var styles = _workbook.Styles;
        if (IsCells)
        {
            var setAll = GetModifyBorderFunc(border => border with
            {
                Left = modify(border.Left, value),
                Top = modify(border.Top, value),
                Right = modify(border.Right, value),
                Bottom = modify(border.Bottom, value),
            }, styles);
            foreach (var (sheetName, area) in Areas)
            {
                if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                    continue;

                ApplyToAll(area, setAll, worksheet);
            }

            return;
        }

        // Change only top and bottom border of a row. The style is used by non-materialized cells
        // in a row and will be used by non-materialized cells in a row. Same applies to columns.
        var setTopAndBottom = GetModifyBorderFunc(border => border with
        {
            Top = modify(border.Top, value),
            Bottom = modify(border.Bottom, value),
        }, styles);
        ModifyRowsBorder(Rows, setTopAndBottom);

        var setLeftAndRight = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Right = modify(border.Right, value),
        }, styles);
        ModifyColumnsBorder(Columns, setLeftAndRight);

        // A normal path for range API object (except XLCells). Set outer border to areas.
        // Don't use UsedAreas, they are for columns/rows. Worksheet doesn't have an outer border.
        var setLeft = GetModifyBorderFunc(border => border with { Left = modify(border.Left, value) }, styles);
        var setTop = GetModifyBorderFunc(border => border with { Top = modify(border.Top, value) }, styles);
        var setRight = GetModifyBorderFunc(border => border with { Right = modify(border.Right, value) }, styles);
        var setBottom = GetModifyBorderFunc(border => border with { Bottom = modify(border.Bottom, value) }, styles);
        foreach (var area in Areas)
        {
            if (!_workbook.TryGetWorksheet(area.Name, out XLWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;

            // Left side
            var left = area.Area.SliceFromLeft(1);
            cellsCollection.ApplyFormatOnAll(left, setLeft, formatResolver.Resolve);

            // Top side
            var top = area.Area.SliceFromTop(1);
            cellsCollection.ApplyFormatOnAll(top, setTop, formatResolver.Resolve);

            // Right side
            var right = area.Area.SliceFromRight(1);
            cellsCollection.ApplyFormatOnAll(right, setRight, formatResolver.Resolve);

            // Bottom side
            var bottom = area.Area.SliceFromBottom(1);
            cellsCollection.ApplyFormatOnAll(bottom, setBottom, formatResolver.Resolve);
        }
    }

    internal void ModifyInnerBorder<TProperty>(Func<XLBorderLine, TProperty, XLBorderLine> modify, TProperty value)
    {
        // Shortcut for XLCells - it has no inner borders
        if (IsCells)
            return;

        var styles = _workbook.Styles;
        ModifyInsideBordersOfRows(styles, modify, value);
        ModifyInsideBordersOfColumns(styles, modify, value);

        var setLeft = GetModifyBorderFunc(border => border with { Left = modify(border.Left, value) }, styles);
        var setTop = GetModifyBorderFunc(border => border with { Top = modify(border.Top, value) }, styles);
        var setRight = GetModifyBorderFunc(border => border with { Right = modify(border.Right, value) }, styles);
        var setBottom = GetModifyBorderFunc(border => border with { Bottom = modify(border.Bottom, value) }, styles);
        foreach (var (sheetName, area) in Areas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;

            // Setting line from both sides is not super useful, but keeps internal state consistent.
            if (area.Width > 1)
            {
                cellsCollection.ApplyFormatOnAll(area.SliceFromLeft(area.Width - 1), setRight, formatResolver.Resolve);
                cellsCollection.ApplyFormatOnAll(area.SliceFromRight(area.Width - 1), setLeft, formatResolver.Resolve);
            }

            if (area.Height > 1)
            {
                cellsCollection.ApplyFormatOnAll(area.SliceFromTop(area.Height - 1), setBottom, formatResolver.Resolve);
                cellsCollection.ApplyFormatOnAll(area.SliceFromBottom(area.Height - 1), setTop, formatResolver.Resolve);
            }
        }
    }

    private static Func<XLCellFormatValue, XLCellFormatValue> GetModifyBorderFunc(Func<XLBorderFormatValue, XLBorderFormatValue> modifyBorder, XLWorkbookStyles styles)
    {
        return format =>
        {
            var modifiedBorder = styles.GetRegisteredBorderFormat(format.Border, modifyBorder);
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with { Border = modifiedBorder });
            return modifiedFormat;
        };
    }

    private void ModifyInsideBordersOfRows<TProperty>(XLWorkbookStyles styles, Func<XLBorderLine, TProperty, XLBorderLine> modify, TProperty value)
    {
        // For a single row, only the left are right border are counted as "inside". The top and bottom border touch the outside.
        var setLeftAndRight = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Right = modify(border.Right, value),
        }, styles);

        // For multi-row rowspan, there are three different patterns:
        // Multi-row rowspan - top row
        var setLeftRightBottom = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-row rowspan - center rows. There isn't a center row in 2-row rowspan
        var setAll = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-row rowspan - bottom row
        var setLeftTopRight = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
        }, styles);

        // Set border for each rowspan
        for (var i = 0; i < Rows.Length; ++i)
        {
            // Find rowspan as a sequence of consecutive rows 
            var startIndex = i;
            var endIndex = i;
            while (endIndex + 1 < Rows.Length && Rows[endIndex + 1].RowNumber == Rows[endIndex].RowNumber + 1)
            {
                endIndex += 1;
            }

            i = endIndex;
            var rowSpanHeight = endIndex - startIndex + 1;
            if (rowSpanHeight > 1)
            {
                ModifyRowsBorder(Rows.AsSpan(startIndex, 1), setLeftRightBottom);
                ModifyRowsBorder(Rows.AsSpan(startIndex + 1, rowSpanHeight - 2), setAll);
                ModifyRowsBorder(Rows.AsSpan(endIndex, 1), setLeftTopRight);
            }
            else
            {
                ModifyRowsBorder(Rows.AsSpan(startIndex, 1), setLeftAndRight);
            }
        }
    }

    private void ModifyRowsBorder(ReadOnlySpan<XLRowArea> rows, Func<XLCellFormatValue, XLCellFormatValue> modifyBorder)
    {
        foreach (var rowArea in rows)
        {
            if (!_workbook.TryGetWorksheet(rowArea.Name, out XLWorksheet worksheet))
                continue;

            // Row style is used by non-materialized cells in a row...
            var row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, modifyBorder, worksheet);

            // ... and materialized cells in a row have format explicitly set.
            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(rowArea.Area.Area, modifyBorder, formatResolver.Resolve);
        }
    }

    private void ModifyInsideBordersOfColumns<TProperty>(XLWorkbookStyles styles, Func<XLBorderLine, TProperty, XLBorderLine> modify, TProperty value)
    {
        // For a single column, only the top are bottom border are counted as "inside". The left and right border touch the outside.
        var setTopAndBottom = GetModifyBorderFunc(border => border with
        {
            Top = modify(border.Top, value),
            Bottom = modify(border.Bottom, value)
        }, styles);

        // For multi-column colspan, there are three different patterns:
        // Multi-column colspan - left column
        var setTopRightBottom = GetModifyBorderFunc(border => border with
        {
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-column colspan - center columns. There isn't a center column in 2-column colspan
        var setAll = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-column colspan - right column
        var setLeftTopBottom = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Set border for each colspan
        for (var i = 0; i < Columns.Length; ++i)
        {
            // Find colspan as a sequence of consecutive columns 
            var startIndex = i;
            var endIndex = i;
            while (endIndex + 1 < Columns.Length && Columns[endIndex + 1].ColumNumber == Columns[endIndex].ColumNumber + 1)
            {
                endIndex += 1;
            }

            i = endIndex;
            var colspanWidth = endIndex - startIndex + 1;
            if (colspanWidth > 1)
            {
                ModifyColumnsBorder(Columns.AsSpan(startIndex, 1), setTopRightBottom);
                ModifyColumnsBorder(Columns.AsSpan(startIndex + 1, colspanWidth - 2), setAll);
                ModifyColumnsBorder(Columns.AsSpan(endIndex, 1), setLeftTopBottom);
            }
            else
            {
                ModifyColumnsBorder(Columns.AsSpan(startIndex, 1), setTopAndBottom);
            }
        }
    }

    private void ModifyColumnsBorder(ReadOnlySpan<XLColumnArea> columns, Func<XLCellFormatValue, XLCellFormatValue> modifyBorder)
    {
        foreach (var columnArea in columns)
        {
            if (!_workbook.TryGetWorksheet(columnArea.Name, out XLWorksheet worksheet))
                continue;

            // Column style is used by non-materialized cells in a column...
            var column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, modifyBorder, worksheet);

            // ... and materialized cells in a column have format explicitly set.
            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(columnArea.Area.Area, modifyBorder, formatResolver.Resolve);
        }
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
