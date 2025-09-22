using System;
using ClosedXML.Excel.Formatting;
using System.Collections.Generic;
using System.Diagnostics;
using ClosedXML.Utils;

namespace ClosedXML.Excel;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XLWorkbookStyles
{
    private readonly BiDictionary<int, string> _numberFormats;

    private readonly BiDictionary<int, XLFontFormatValue> _fontFormats;

    private readonly BiDictionary<int, XLFillFormatValue> _fillFormats;

    private readonly BiDictionary<int, XLBorderFormatValue> _borderFormats;

    /// <summary>
    /// The key is XfId, the value is cell format.
    /// </summary>
    private readonly BiDictionary<int, XLCellFormatValue> _cellFormats;

    /// <summary>
    /// The key is cellStyleXfId, the value is cell style.
    /// </summary>
    private readonly BiDictionary<StyleId, XLCellStyleValue> _cellStyles;

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
        _numberFormats = new BiDictionary<int, string>();
        _fontFormats = new BiDictionary<int, XLFontFormatValue>();
        _fillFormats = new BiDictionary<int, XLFillFormatValue>();
        _borderFormats = new BiDictionary<int, XLBorderFormatValue>();
        _cellFormats = new BiDictionary<int, XLCellFormatValue>();
        _cellStyles = new BiDictionary<StyleId, XLCellStyleValue>();
        _differentialFormats = new BiDictionary<int, XLDxfValue>();
        _tableStyles = new Dictionary<string, XLTableTheme>(XLHelper.NameComparer);
        _pivotStyles = new Dictionary<string, XLPivotTableStyle>(XLHelper.NameComparer);
    }

    internal IReadOnlyBiDictionary<int, string> NumberFormats => _numberFormats;

    internal IReadOnlyBiDictionary<int, XLFontFormatValue> Fonts => _fontFormats;

    internal IReadOnlyBiDictionary<int, XLFillFormatValue> Fills => _fillFormats;

    internal IReadOnlyBiDictionary<int, XLBorderFormatValue> Borders => _borderFormats;

    internal IReadOnlyBiDictionary<int, XLCellFormatValue> CellFormats => _cellFormats;

    internal IReadOnlyBiDictionary<StyleId, XLCellStyleValue> CellStyles => _cellStyles;

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

    /// <summary>
    /// A default format values used when format doesn't have a value in a property. All props in
    /// the default format must have a value. The default is set on load and not changed later.
    /// Nearly all props are equivalent of "zero", except things that can't be like that, e.g. font
    /// name or font size.
    /// </summary>
    // TODO: Make private and use GetDefaultFormat
    internal XLCellFormatValue DefaultFormat { get; set; } = new()
    {
        Font = new XLFontFormatValue
        {
            Name = "Calibri",
            Charset = XLFontCharSet.Ansi,
            Family = XLFontFamilyNumberingValues.NotApplicable,
            Bold = false,
            Italic = false,
            Strikethrough = false,
            Outline = false,
            Shadow = false,
            Condense = false,
            Extend = false,
            Color = XLColor.FromArgb(0x00000000),
            Size = XLFontSize.FromPoints(11),
            Underline = XLFontUnderlineValues.None,
            VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
            Scheme = XLFontScheme.None
        },
        // TODO: Add all default values, not just font
        NumberFormat = null,
        Alignment = null,
        Protection = null,
        Fill = null,
        Border = null,
        CellStyleId = null,
        IncludeQuotePrefix = false,
        PivotButton = false,
        CustomFormat = CellFormatComponents.None
    };

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

    internal void AddCellStyle(int cellStyleXfId, XLCellStyleValue cellStyle)
    {
        _cellStyles.Add(cellStyleXfId, cellStyle);
    }

    internal void AddDifferentialFormat(XLDxfValue dxf)
    {
        _differentialFormats.Add(_differentialFormats.Count, dxf);
    }

    internal void AddTableStyle(XLTableTheme tableStyle)
    {
        _tableStyles.Add(tableStyle.Name, tableStyle);
    }

    internal void AddPivotStyle(XLPivotTableStyle pivotStyle)
    {
        _pivotStyles.Add(pivotStyle.Name, pivotStyle);
    }

    internal void SetIndexedColors(List<uint> indexedColors)
    {
        _indexedColorsArgb = indexedColors;
    }

    internal void SetMruColors(List<XLColor> mruColors)
    {
        _mruColors = mruColors;
    }

    /// <summary>
    /// Get a font format that is stored in the internal structures of the styles class. The font
    /// format is created by modification of existing font format. This is essential for saving,
    /// all formats must be registered in the styles class. 
    /// </summary>
    internal XLFontFormatValue GetRegisteredFontFormat(XLFontFormatValue original, Func<XLFontFormatValue, XLFontFormatValue> modify)
    {
        var modified = modify(original);
        if (_fontFormats.TryGetValue(modified, out var existingFont))
            return existingFont;

        AddFontFormat(modified);
        return modified;
    }

    internal XLCellFormatValue GetRegisteredCellFormat(XLCellFormatValue original, Func<XLCellFormatValue, XLCellFormatValue> modify)
    {
        var modified = modify(original);
        if (_cellFormats.TryGetValue(modified, out var existing))
            return existing;

        AddFormat(modified);
        return modified;
    }

    /// <summary>
    /// Get a differential format that is stored in the internal structures of the styles class.
    /// The differential format is created by modification of existing dxf format. This is
    /// essential for saving, all formats must be registered in the styles class. 
    /// </summary>
    internal XLDxfValue GetRegisteredDxFormat(XLDxfValue original, Func<XLDxfValue, XLDxfValue> modify)
    {
        var modified = modify(original);
        if (_differentialFormats.TryGetValue(modified, out var existingDxf))
            return existingDxf;

        AddDifferentialFormat(modified);
        return modified;
    }

    internal T GetDefaultFormat<T>(Func<XLCellFormatValue, T?> selector)
    {
        return selector(DefaultFormat) ?? throw new UnreachableException("Default value doesn't contain a format property.");
    }
}
