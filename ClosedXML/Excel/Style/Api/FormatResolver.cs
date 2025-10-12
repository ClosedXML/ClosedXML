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
    private readonly XLWorksheet _worksheet;
    private readonly XLColumnsCollection _columns;
    private readonly XLRowsCollection _rows;

    public FormatResolver(XLWorksheet worksheet)
    {
        _defaultFormat = worksheet.Workbook.Styles.DefaultFormat;
        _worksheet = worksheet;
        _columns = worksheet.Internals.ColumnsCollection;
        _rows = worksheet.Internals.RowsCollection;
    }

    /// <summary>
    /// Resolve a style of a point in a worksheet according to format hierarchy.
    /// </summary>
    /// <param name="point">Point for which to resolve the style.</param>
    /// <returns>A format that is already registered in the styles.</returns>
    public XLCellFormatValue Resolve(XLSheetPoint point)
    {
        // TODO: Add resolve from worksheet
        if (_rows.TryGetValue(point.Row, out var row) &&
            row.FormatValue is { } rowFormat)
        {
            return rowFormat;
        }

        if (_columns.TryGetValue(point.Column, out var column) &&
            column.FormatValue is { } columnFormat)
        {
            return columnFormat;
        }

        if (_worksheet.FormatValue is { } worksheetFormat)
            return worksheetFormat;

        return _defaultFormat;
    }
}
