using System;
using System.Collections.Generic;
using System.Diagnostics;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// Unlike the <see cref="XLStyle"/>, the <see cref="XLCellFormat"/> one modifies formatting
/// in a <see cref="XLWorkbookStyles"/>.
/// </summary>
internal class XLCellFormat
{
    /// <summary>
    /// A workbook components with a format that should be updated when user uses API to change
    /// a format (e.g. a row). Container doesn't have to initially have a value, e.g. it's a cell
    /// or a row without a format (thus inherits normal).
    /// </summary>
    private readonly IXLFormatContainer _container;

    /// <summary>
    /// Cell areas in a workbook that should be updated when format is changed, e.g. when we have
    /// a format API object for a row container, the area are all cells of the row. It must be
    /// an area, so we can satisfy the <see cref="IXLBorder.OutsideBorder"/> and
    /// <see cref="IXLBorder.InsideBorder"/> property setters.
    /// </summary>
    private readonly IReadOnlyList<XLBookArea> _areas = Array.Empty<XLBookArea>();

    private readonly XLWorkbookStyles _styles;

    /// <summary>
    /// Row and column formats. They don't do anything for setting the value, but are useful when
    /// asking for a value of a non-existent or non-styles cell.
    /// </summary>
    /// <remarks>
    /// When a cell has no explicit style and column and row cross, the row style
    /// wins (i.e. when both column and row specify color, the row one is displayed when cell
    /// doesn't have a style or doesn't exist). Therefore row is asked before column.
    /// </remarks>
    private readonly IXLFormatContainer? _row;
    private readonly IXLFormatContainer? _column;

    /// <summary>
    /// A normal style of a workbook. Should have all values.
    /// </summary>
    private readonly IXLFormatContainer _normal;

    private XLCellFormat(IXLFormatContainer container, IXLFormatContainer? row, IXLFormatContainer? column, IXLFormatContainer normal, XLWorkbookStyles styles)
    {
        _container = container;
        _row = row;
        _column = column;
        _normal = normal;
        _styles = styles ?? throw new ArgumentNullException(nameof(styles));
    }

    internal XLFontCellFormat Font => new(this);

    internal static XLCellFormat ForCell(XLWorkbookStyles styles, IXLFormatContainer cell, IXLFormatContainer? row, IXLFormatContainer? column, IXLFormatContainer normal)
    {
        return new XLCellFormat(cell, row, column, normal, styles);
    }

    internal static XLCellFormat ForRow(XLWorkbookStyles styles, IXLFormatContainer row, IXLFormatContainer normal)
    {
        return new XLCellFormat(row, null, null, normal, styles);
    }

    internal static XLCellFormat ForColumn(XLWorkbookStyles styles, IXLFormatContainer column, IXLFormatContainer normal)
    {
        return new XLCellFormat(column, null, null, normal, styles);
    }

    internal T Resolve<T>(Func<XLCellFormatValue, T?> selector)
        where T : struct
    {
        var defaultFormat = _styles.DefaultFormat;
        var format = GetFormatOrInherited(defaultFormat);
        return GetFormatProperty(selector, format, defaultFormat);
    }

    private XLCellFormatValue GetFormatOrInherited(XLCellFormatValue defaultFormat)
    {
        // Get the format in hierarchy that is closest to the actual container.
        if (_container.FormatValue is { } containerFormat)
            return containerFormat;

        if (_row?.FormatValue is { } rowFormat)
            return rowFormat;

        if (_column?.FormatValue is { } columnFormat)
            return columnFormat;

        if (_normal.FormatValue is { } normalFormat)
            return normalFormat;

        // We should never get here, but if workbook doesn't specify normal style (technically not
        // required by the spec), let's go with the default format.
        return defaultFormat;
    }

    private T GetFormatProperty<T>(Func<XLCellFormatValue, T?> selector, XLCellFormatValue format, XLCellFormatValue defaultFormat)
        where T : struct
    {
        var formatPropertyValue = selector(format);
        if (formatPropertyValue is not null)
            return formatPropertyValue.Value;

        var defaultPropertyValue = selector(defaultFormat);
        if (defaultPropertyValue is not null)
            return defaultPropertyValue.Value;

        throw new UnreachableException("Default format is missing a value.");
    }

    internal void ModifyFont<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        var format = _container.FormatValue ?? GetFormatOrInherited(_styles.DefaultFormat);
        var font = format.Font ?? _styles.GetDefaultFormat(static x => x.Font);
        var modifiedFont = _styles.GetRegisteredFontFormat(font, f => modifyFont(f, value));
        var modifiedFormat = _styles.GetRegisteredCellFormat(format, f => f with { Font = modifiedFont });
        _container.FormatValue = modifiedFormat;
    }
}
