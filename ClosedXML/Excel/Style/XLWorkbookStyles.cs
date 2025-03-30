using ClosedXML.Excel.Formatting;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XLWorkbookStyles
{
    /// <summary>
    /// The index is XfId, the value is formatting record.
    /// </summary>
    private List<XLCellFormat> _masterFormats;

    private readonly List<XLFontFormat> _fontFormats;

    internal XLWorkbookStyles(List<XLCellFormat> masterFormats, List<XLFontFormat> fontFormats)
    {
        _masterFormats = masterFormats;
        _fontFormats = fontFormats;
    }

    internal IReadOnlyList<XLFontFormat> FontFormats => _fontFormats;
}
