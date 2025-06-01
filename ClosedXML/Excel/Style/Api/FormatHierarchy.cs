using System;
using System.Collections.Generic;
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
    /// Containers that are modified by the API changes.
    /// </summary>
    private readonly IXLFormatContainer[] _containers;

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
    /// Sheet format is a ClosedXML fiction. OOXML doesn't have a style for a sheet. Think of it as
    /// all columns without a specified format. The xml tag <![CDATA[<col/>]]> has attributes
    /// <c>min</c> (column) and <c>max</c> (column), so we can easily specify formatting for
    /// all columns from start to end (or just ones that don't have explicit format).
    /// </summary>
    private readonly IXLFormatContainer? _sheet;

    /// <summary>
    /// A normal style of a workbook. Should have all values.
    /// </summary>
    private readonly IXLFormatContainer _normal;

    public FormatHierarchy(IXLFormatContainer[] containers, IXLFormatContainer? row, IXLFormatContainer? column, IXLFormatContainer? sheet, IXLFormatContainer normal)
    {
        _containers = containers ?? throw new ArgumentNullException(nameof(containers));
        _row = row;
        _column = column;
        _sheet = sheet;
        _normal = normal;
    }

    public T Resolve<T>(Func<XLCellFormatValue, T?> getFormatValue, T existingFormatMissingValueDefault)
        where T : struct
    {
        var container = _containers[0];
        var format = container.FormatValue;
        if (format is not null)
        {
            return getFormatValue(format) ?? existingFormatMissingValueDefault;
        }

        // Container doesn't even have a format, e.g. it's a style-less cell or a row. 
        throw new NotImplementedException();
    }

    public T ResolveWithNormalFallback<T>(Func<XLCellFormatValue, T?> getFormatValue)
        where T : struct
    {
        var normalFormat = _normal.FormatValue ?? throw new InvalidOperationException("Normal style doesn't have format");
        var normalValue = getFormatValue(normalFormat) ?? throw new InvalidOperationException("Normal style should have everything.");
        return Resolve(getFormatValue, normalValue);
    }
}
