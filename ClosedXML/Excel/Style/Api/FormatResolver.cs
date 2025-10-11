using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// A class to resolve a format of a cell in a worksheet, if the cell doesn't already has format.
/// It basically goes up the levels and tries to find the "closest" component with a non-empty
/// format. This is a short lived class to be used modifying format of one sheet area.
/// </summary>
internal class FormatResolver
{
    private readonly XLCellFormatValue _defaultFormat;
    private readonly XLColumnsCollection _columns;

    public FormatResolver(XLWorksheet worksheet)
    {
        _defaultFormat = worksheet.Workbook.Styles.DefaultFormat;
        _columns = worksheet.Internals.ColumnsCollection;
    }

    /// <summary>
    /// Resolve a style of a point in a worksheet according to format hierarchy.
    /// </summary>
    /// <param name="point">Point for which to resolve the style.</param>
    /// <returns>A format that is already registered in the styles.</returns>
    public XLCellFormatValue Resolve(XLSheetPoint point)
    {
        // TODO: Add resolve from worksheet and rows
        if (_columns.TryGetValue(point.Column, out var column) &&
            column.FormatValue is { } columnFormat)
        {
            return columnFormat;
        }

        return _defaultFormat;
    }
}
