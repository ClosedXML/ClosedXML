using System;
using ClosedXML.Excel.Formatting;
using System.Collections.Generic;
using System.Globalization;

namespace ClosedXML.Excel;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XLWorkbookStyles
{
    private readonly Dictionary<int, string> _numberFormats;

    private readonly Dictionary<int, XLFontFormat> _fontFormats;

    private readonly Dictionary<int, XLFillFormat> _fillFormats;

    private readonly Dictionary<int, XLBorderFormat> _borderFormats;

    /// <summary>
    /// The key is XfId, the value is cell format.
    /// </summary>
    private readonly Dictionary<int, XLCellFormat> _cellFormats;

    /// <summary>
    /// The key is cellStyleXfId, the value is cell style.
    /// </summary>
    private readonly Dictionary<int, XLCellStyle> _cellStyles;

    internal XLWorkbookStyles()
    {
        _numberFormats = new Dictionary<int, string>();
        _fontFormats = new Dictionary<int, XLFontFormat>();
        _fillFormats = new Dictionary<int, XLFillFormat>();
        _borderFormats = new Dictionary<int, XLBorderFormat>();
        _cellFormats = new Dictionary<int, XLCellFormat>();
        _cellStyles = new Dictionary<int, XLCellStyle>();
    }

    internal IReadOnlyDictionary<int, string> NumberFormats => _numberFormats;

    internal IReadOnlyDictionary<int, XLFontFormat> Fonts=> _fontFormats;

    internal IReadOnlyDictionary<int, XLFillFormat> Fills => _fillFormats;

    internal IReadOnlyDictionary<int, XLBorderFormat> Borders => _borderFormats;

    internal IReadOnlyDictionary<int, XLCellFormat> CellFormats => _cellFormats;

    internal IReadOnlyDictionary<int, XLCellStyle> CellStyles => _cellStyles;

    internal XLNumberFormatValue GetNumberFormat(int numberFormatId)
    {
        var xlNumberFormat = new XLNumberFormatKey
        {
            NumberFormatId = numberFormatId,
            Format = _numberFormats[numberFormatId]
        };
        return XLNumberFormatValue.FromKey(ref xlNumberFormat);
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

    internal void AddFormat(XLCellFormat cellFormat)
    {
        var xfId = _cellFormats.Count;
        _cellFormats.Add(xfId, cellFormat);
    }

    public void AddCellStyle(int cellStyleXfId, XLCellStyle cellStyle)
    {
        _cellStyles.Add(cellStyleXfId, cellStyle);
    }
}
