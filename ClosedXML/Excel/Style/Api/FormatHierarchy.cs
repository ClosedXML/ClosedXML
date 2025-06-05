using System;
using System.Collections.Generic;
using System.Diagnostics;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// Format hierarchy is used to resolve a format properties. It has three main properties:
/// <list type="bullet">
///   <item>
///     <term>Containers</term>
///     <description>
///     A workbook components with a format that should be updated when user uses API to change
///     a format (e.g. a row). Container doesn't have to initially have a value, e.g. it's a cell
///     or a row without a format (thus inherits normal).
///     </description>
///   </item>
///   <item>
///     <term>Areas</term>
///     <description>
///     Cell areas in a workbook that should be updated when format is changed, e.g. when we have
///     a format API object for a row container, the area are all cells of the row. It must be
///     an area, so we can satisfy the <see cref="IXLBorder.OutsideBorder"/> and
///     <see cref="IXLBorder.InsideBorder"/> property setters.
///     </description>
///   </item>
///   <item>
///     <term>Format hierarchy</term>
///     <description>
///     It contains info about formats of format objects that are higher in a hierarchy. For
///     a cell, higher objects are row, column, sheet and normal style. For a column, higher object
///     is a sheet and then normal style. This info is used to determine format for containers
///     without a format (don't create format unless necessary).
///     </description>
///   </item>
/// </list>
/// </summary>
internal readonly struct FormatHierarchy
{
    /// <summary>
    /// Container that is modified by the API changes.
    /// </summary>
    private readonly IXLFormatContainer _container;

    /// <summary>
    /// Cell areas that should have format updated upon a format change.
    /// </summary>
    private readonly IReadOnlyList<XLBookArea> _areas = Array.Empty<XLBookArea>();

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

    private FormatHierarchy(IXLFormatContainer container, IXLFormatContainer? row, IXLFormatContainer? column, IXLFormatContainer normal)
    {
        _container = container ?? throw new ArgumentNullException(nameof(container));
        _row = row;
        _column = column;
        _normal = normal;
    }

    public static FormatHierarchy ForCell(IXLFormatContainer cell, IXLFormatContainer? row, IXLFormatContainer? column, IXLFormatContainer normal)
    {
        return new FormatHierarchy(cell, row, column, normal);
    }

    public static FormatHierarchy ForRow(IXLFormatContainer row, IXLFormatContainer normal)
    {
        return new FormatHierarchy(row, null, null, normal);
    }

    public static FormatHierarchy ForColumn(IXLFormatContainer column, IXLFormatContainer normal)
    {
        return new FormatHierarchy(column, null, null, normal);
    }

    public T Resolve<T>(Func<XLCellFormatValue, T?> getFormatValue, XLCellFormatValue defaultFormat)
        where T : struct
    {
        var format = GetNearestFormat(defaultFormat);
        var formatPropertyValue = getFormatValue(format);
        if (formatPropertyValue is not null)
            return formatPropertyValue.Value;

        var defaultPropertyValue = getFormatValue(defaultFormat);
        if (defaultPropertyValue is not null)
            return defaultPropertyValue.Value;

        throw new UnreachableException("Default format is missing a value.");
    }

    private XLCellFormatValue GetNearestFormat(XLCellFormatValue defaultFormat)
    {
        // Get the format in hierarchy that is closest to the actual container.
        if (_container.FormatValue is { } containerFormat)
            return containerFormat;

        if (_row?.FormatValue is { } rowFormat)
            return rowFormat;

        if (_column?.FormatValue is { } columnFormat)
            return columnFormat;

        if (_normal?.FormatValue is { } normalFormat)
            return normalFormat;

        // We should never get here, but if workbook doesn't specify normal style (technically not
        // required by the spec), let's go with the default format.
        return defaultFormat;
    }
}
