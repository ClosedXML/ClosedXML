using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// An interface used by object that use differential formatting, e.g. conditional formatting or
/// pivot tables.
/// </summary>
internal interface IXLDxfContainer
{
    /// <summary>
    /// The differential format value.
    /// </summary>
    /// <remarks>
    /// The value is optional, because attribute in XML is optional.
    /// </remarks>
    XLDxfValue? FormatValue { get; set; }
}
