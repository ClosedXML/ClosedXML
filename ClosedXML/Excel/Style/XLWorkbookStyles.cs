using ClosedXML.Excel.Formatting;
using System.Collections.Generic;
using ClosedXML.Utils;

namespace ClosedXML.Excel;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XLWorkbookStyles
{
    private readonly Dictionary<int, string> _numberFormats;

    private readonly BiDictionary<int, XLFontFormatValue> _fontFormats;

    private readonly Dictionary<int, XLFillFormatValue> _fillFormats;

    private readonly Dictionary<int, XLBorderFormatValue> _borderFormats;

    /// <summary>
    /// The key is XfId, the value is cell format.
    /// </summary>
    private readonly Dictionary<int, XLCellFormatValue> _cellFormats;

    /// <summary>
    /// The key is cellStyleXfId, the value is cell style.
    /// </summary>
    private readonly Dictionary<int, XLCellStyleValue> _cellStyles;

    private readonly BiDictionary<int, XLDxfValue> _differentialFormats;

    /// <summary>
    /// Key is a table style name, value is a table style.
    /// </summary>
    private readonly Dictionary<string, XLTableTheme> _tableStyles;

    /// <summary>
    /// Key is a pivot table style name, value is a pivot table style.
    /// </summary>
    private readonly Dictionary<string, XLPivotTableStyle> _pivotStyles;

    private List<uint>? _indexedColorsArgb;

    private List<XLColor> _mruColors = new();

    internal XLWorkbookStyles()
    {
        _numberFormats = new Dictionary<int, string>();
        _fontFormats = new BiDictionary<int, XLFontFormatValue>();
        _fillFormats = new Dictionary<int, XLFillFormatValue>();
        _borderFormats = new Dictionary<int, XLBorderFormatValue>();
        _cellFormats = new Dictionary<int, XLCellFormatValue>();
        _cellStyles = new Dictionary<int, XLCellStyleValue>();
        _differentialFormats = new BiDictionary<int, XLDxfValue>();
        _tableStyles = new Dictionary<string, XLTableTheme>(XLHelper.NameComparer);
        _pivotStyles = new Dictionary<string, XLPivotTableStyle>(XLHelper.NameComparer);
    }

    internal IReadOnlyDictionary<int, string> NumberFormats => _numberFormats;

    internal IReadOnlyDictionary<int, XLFontFormatValue> Fonts => _fontFormats.KeyToValue;

    internal IReadOnlyDictionary<int, XLFillFormatValue> Fills => _fillFormats;

    internal IReadOnlyDictionary<int, XLBorderFormatValue> Borders => _borderFormats;

    internal IReadOnlyDictionary<int, XLCellFormatValue> CellFormats => _cellFormats;

    internal IReadOnlyDictionary<int, XLCellStyleValue> CellStyles => _cellStyles;

    internal IReadOnlyDictionary<int, XLDxfValue> DifferentialFormats => _differentialFormats.KeyToValue;

    internal IReadOnlyDictionary<string, XLTableTheme> TableStyles => _tableStyles;

    internal IReadOnlyDictionary<string, XLPivotTableStyle> PivotStyles => _pivotStyles;

    /// <summary>
    /// Name of a table style that should be used for newly added tables. It's not used for tables
    /// without a specified style.
    /// </summary>
    internal string? DefaultTableStyle { get; set; }

    /// <summary>
    /// Name of a pivot style that should be used for newly added pivot tables. It's not used for
    /// pivot tables without a specified style.
    /// </summary>
    internal string? DefaultPivotStyle { get; set; }

    /// <summary>
    /// Some workbooks use indexed colors that are not in the standard <see cref="XLColor.IndexedColors"/>,
    /// but have their own list of indexed colors. Legacy feature, do not expose. If the value is null, use
    /// predefined indexed colors.
    /// </summary>
    internal IReadOnlyList<uint>? IndexedColorsArgb => _indexedColorsArgb;

    /// <summary>
    /// Most recently used colors are colors <em>hand-picked</em> by a user that are displayed in
    /// the color picker dialogue. The standard and theme colors are always offered in the color
    /// picker, so they are (in general) not added to the MRU color list.
    /// </summary>
    internal IReadOnlyList<XLColor> MruColors => _mruColors;

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

    internal void AddFontFormat(XLFontFormatValue fontFormat)
    {
        _fontFormats.Add(_fontFormats.Count, fontFormat);
    }

    internal void AddFillFormat(XLFillFormatValue fillFormat)
    {
        _fillFormats.Add(_fillFormats.Count, fillFormat);
    }

    internal void AddBorderFormat(XLBorderFormatValue borderFormat)
    {
        _borderFormats.Add(_borderFormats.Count, borderFormat);
    }

    internal void AddFormat(XLCellFormatValue cellFormat)
    {
        var xfId = _cellFormats.Count;
        _cellFormats.Add(xfId, cellFormat);
    }

    public void AddCellStyle(int cellStyleXfId, XLCellStyleValue cellStyle)
    {
        _cellStyles.Add(cellStyleXfId, cellStyle);
    }

    public void AddDifferentialFormat(XLDxfValue dxf)
    {
        _differentialFormats.Add(_differentialFormats.Count, dxf);
    }

    public void AddTableStyle(XLTableTheme tableStyle)
    {
        _tableStyles.Add(tableStyle.Name, tableStyle);
    }

    public void AddPivotStyle(XLPivotTableStyle pivotStyle)
    {
        _pivotStyles.Add(pivotStyle.Name, pivotStyle);
    }

    public void SetIndexedColors(List<uint> indexedColors)
    {
        _indexedColorsArgb = indexedColors;
    }

    public void SetMruColors(List<XLColor> mruColors)
    {
        _mruColors = mruColors;
    }
}
