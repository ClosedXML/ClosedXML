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
    private readonly Dictionary<int, XLCellFormat> _masterFormats;

    private readonly Dictionary<int, XLFontFormat> _fontFormats;

    internal XLWorkbookStyles()
    {
        _masterFormats = new Dictionary<int, XLCellFormat>();
        _fontFormats = new Dictionary<int, XLFontFormat>();
    }

    internal XLStyleKey ApplyFontFormat(int fontId, ref XLStyleKey xlStyle)
    {
        var fontFormat = _fontFormats[fontId];
        var xlFont = fontFormat.ApplyTo(xlStyle.Font);
        return xlStyle with { Font = xlFont };
    }

    internal void AddFontFormat(XLFontFormat fontFormat)
    {
        _fontFormats.Add(_fontFormats.Count, fontFormat);
    }

    internal void AddFormat(uint? fontId)
    {
        var xfId = _masterFormats.Count;
        XLFontFormat? font = fontId is not null ? _fontFormats[checked((int)fontId)] : null;
        _masterFormats.Add(xfId, new XLCellFormat
        {
            Font = font
        });
    }
}
