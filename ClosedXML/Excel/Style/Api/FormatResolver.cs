using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// A class to resolve a format of a cell in a worksheet, if the cell doesn't already has format.
/// It basically goes up the levels and tries to find the "closest" component with a non-empty
/// format. This is a short lived class to be used modifying format of one sheet area.
/// </summary>
internal class FormatResolver
{
    private readonly XLCellFormatValue _workbookFormat;

    public FormatResolver(XLWorksheet worksheet)
    {
        _workbookFormat = worksheet.Workbook.Styles.DefaultFormat;
    }

    /// <summary>
    /// Resolve a style of a point in a worksheet according to format hierarchy.
    /// </summary>
    /// <param name="point">Point for which to resolve the style.</param>
    /// <returns>A format that is already registered in the styles.</returns>
    public XLCellFormatValue Resolve(XLSheetPoint point)
    {
        // TODO: Add resolve from worksheet, columns and rows
        return _workbookFormat;
    }
}
