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

    private readonly Dictionary<int, XLBorderFormat> _borderFormats;

    internal XLWorkbookStyles()
    {
        _masterFormats = new Dictionary<int, XLCellFormat>();
        _fontFormats = new Dictionary<int, XLFontFormat>();
        _borderFormats = new Dictionary<int, XLBorderFormat>();
    }

    internal XLStyleKey ApplyFontFormat(int fontId, ref XLStyleKey styleKey)
    {
        var fontFormat = _fontFormats[fontId];
        var fontKey = fontFormat.ApplyTo(styleKey.Font);
        return styleKey with { Font = fontKey };
    }

    internal XLStyleKey ApplyBorderFormat(int borderId, ref XLStyleKey styleKey)
    {
        var borderFormat = _borderFormats[borderId];
        var borderKey = borderFormat.ApplyTo(styleKey.Border);
        return styleKey with { Border = borderKey };
    }

    internal void AddFontFormat(XLFontFormat fontFormat)
    {
        _fontFormats.Add(_fontFormats.Count, fontFormat);
    }

    internal void AddBorderFormat(XLBorderFormat borderFormat)
    {
        _borderFormats.Add(_borderFormats.Count, borderFormat);
    }

    internal void AddFormat(uint? fontId, uint? borderId)
    {
        var xfId = _masterFormats.Count;
        XLFontFormat? font = fontId is not null ? _fontFormats[checked((int)fontId)] : null;
        XLBorderFormat? border = borderId is not null ? _borderFormats[checked((int)borderId)] : null;
        _masterFormats.Add(xfId, new XLCellFormat
        {
            Font = font,
            Border = border,
        });
    }
}
