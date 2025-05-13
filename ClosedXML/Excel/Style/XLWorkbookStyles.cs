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
    private readonly Dictionary<int, XLCellFormat> _cellFormats;

    private readonly Dictionary<int, string> _numberFormats;

    private readonly Dictionary<int, XLFontFormat> _fontFormats;

    private readonly Dictionary<int, XLFillFormat> _fillFormats;

    private readonly Dictionary<int, XLBorderFormat> _borderFormats;

    internal XLWorkbookStyles()
    {
        _cellFormats = new Dictionary<int, XLCellFormat>();
        _numberFormats = new Dictionary<int, string>();
        _fontFormats = new Dictionary<int, XLFontFormat>();
        _fillFormats = new Dictionary<int, XLFillFormat>();
        _borderFormats = new Dictionary<int, XLBorderFormat>();
    }

    internal IReadOnlyDictionary<int, string> NumberFormats => _numberFormats;

    internal IReadOnlyDictionary<int, XLFontFormat> Fonts=> _fontFormats;

    internal IReadOnlyDictionary<int, XLFillFormat> Fills => _fillFormats;

    internal IReadOnlyDictionary<int, XLBorderFormat> Borders => _borderFormats;

    internal IReadOnlyDictionary<int, XLCellFormat> CellFormats => _cellFormats;

    internal XLStyleKey ApplyNumberFormat(int numberFormatId, ref XLStyleKey styleKey)
    {
        // Unlike other aspects of formatting, number format is skipped when numFmtId is not found
        if (!_numberFormats.TryGetValue(numberFormatId, out var formatCode))
            return styleKey;

        var numberFormat = new XLNumberFormatKey
        {
            NumberFormatId = numberFormatId,
            Format = formatCode
        };
        return styleKey with { NumberFormat = numberFormat };
    }

    internal XLNumberFormatValue GetNumberFormat(int numberFormatId)
    {
        var xlNumberFormat = new XLNumberFormatKey
        {
            NumberFormatId = numberFormatId,
            Format = _numberFormats[numberFormatId]
        };
        return XLNumberFormatValue.FromKey(ref xlNumberFormat);
    }

    internal XLStyleKey ApplyFontFormat(int fontId, ref XLStyleKey styleKey)
    {
        var fontFormat = _fontFormats[fontId];
        var fontKey = fontFormat.ApplyTo(styleKey.Font);
        return styleKey with { Font = fontKey };
    }

    internal XLStyleKey ApplyPatternFormat(int fillId, ref XLStyleKey styleKey)
    {
        var fillFormat = _fillFormats[fillId];
        var fillKey = fillFormat.ApplyTo(styleKey.Fill);
        return styleKey with { Fill = fillKey };
    }

    internal XLStyleKey ApplyBorderFormat(int borderId, ref XLStyleKey styleKey)
    {
        var borderFormat = _borderFormats[borderId];
        var borderKey = borderFormat.ApplyTo(styleKey.Border);
        return styleKey with { Border = borderKey };
    }

    internal void AddNumberFormat(int numFmtId, string formatCode)
    {
        _numberFormats.Add(numFmtId, formatCode);
    }

    internal void AddFontFormat(XLFontFormat fontFormat)
    {
        _fontFormats.Add(_fontFormats.Count, fontFormat);
    }

    internal void AddFillFormat(XLFillFormat fillFormat)
    {
        _fillFormats.Add(_fillFormats.Count, fillFormat);
    }

    internal void AddBorderFormat(XLBorderFormat borderFormat)
    {
        _borderFormats.Add(_borderFormats.Count, borderFormat);
    }

    internal void AddFormat(XLAlignmentFormat? alignment, XLProtectionFormat? protection, XLFontFormat? font, XLFillFormat? fill, XLBorderFormat? border)
    {
        var xfId = _cellFormats.Count;
        _cellFormats.Add(xfId, new XLCellFormat
        {
            Alignment = alignment,
            Protection = protection,
            Font = font,
            Fill = fill,
            Border = border,
        });
    }
}
