using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// An interface for objects that have a cell format, for example cell or a row. This is used by
/// the API to modify the cell formatting. Cell formats are immutable, so they can only be fully
/// replaced, never modified. The set value must be registered in the <see cref="XLWorkbookStyles"/>.
/// </summary>
internal interface IXLFormatContainer
{
    /// <summary>
    /// The format of a container can be <c>null</c>, e.g. a row that doesn't have format. In that
    /// situation, we have a container (e.g. a row), but no format.
    /// </summary>
    XLCellFormatValue? FormatValue { get; set; }
}
