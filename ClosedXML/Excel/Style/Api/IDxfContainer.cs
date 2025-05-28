using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// An interface for objects that have a differential format, for example conditional formatting or
/// a table. This is used by the API to modify the differential formatting. Dxfs are immutable, so
/// they can only be fully replaced, never modified. The set value must be registered
/// in the <see cref="XLWorkbookStyles"/>.
/// </summary>
internal interface IDxfContainer
{
    XLDxfValue FormatValue { get; set; }
}
