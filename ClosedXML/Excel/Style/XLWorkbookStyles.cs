using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel.Formatting;
using ClosedXML.Utils;

namespace ClosedXML.Excel;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XLWorkbookStyles
{
    /// <summary>
    /// First user-defined numFmtId.
    /// </summary>
    public const int FirstUserDefinedNumberFormatIndex = 164;

    private readonly BiDictionary<int, XLNumberFormat> _numberFormats;

    private readonly BiDictionary<int, XLFontFormatValue> _fontFormats;

    private readonly BiDictionary<int, XLFillFormatValue> _fillFormats;

    private readonly BiDictionary<int, XLBorderFormatValue> _borderFormats;

    private readonly BiDictionary<int, XLAlignmentFormatValue> _alignmentFormats;

    private readonly BiDictionary<int, XLProtectionFormatValue> _protectionFormats;

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

    /// <summary>
    /// A normal style that is used for newly create workbooks or loaded workbooks without a normal style.
    /// </summary>
    internal readonly XLCellStyleValue DefaultNormalStyle = new()
    {
        Name = "Normal",
        BuiltInStyle = BuiltInStyleValues.Normal,
        Hidden = false,
        Alignment = new XLAlignmentFormatValue
        {
            Horizontal = XLAlignmentHorizontalValues.General,
            Vertical = XLAlignmentVerticalValues.Bottom,
            TextRotation = TextRotation.None,
            WrapText = false,
            Indent = 0,
            RelativeIndent = 0,
            JustifyLastLine = false,
            ShrinkToFit = false,
            ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent
        },
        Protection = new XLProtectionFormatValue
        {
            Locked = true,
            Hidden = false,
        },
        NumberFormat = XLPredefinedFormat.FormatCodes[XLPredefinedFormat.General],
        Font = new XLFontFormatValue
        {
            Name = "Calibri",
            Charset = XLFontCharSet.Ansi,
            Family = XLFontFamilyNumberingValues.Swiss,
            Bold = false,
            Italic = false,
            Strikethrough = false,
            Outline = false,
            Shadow = false,
            Condense = false,
            Extend = false,
            Color = XLColor.Black,
            Size = XLFontSize.FromPoints(11),
            Underline = XLFontUnderlineValues.None,
            VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
            Scheme = XLFontScheme.None
        },
        Fill = XLFillFormatValue.None,
        Border = XLBorderFormatValue.None,
        IncludedComponents = CellFormatComponents.All
    };

    /// <summary>
    /// A cell format that is used when an element doesn't explicitly define formatting. This
    /// format must be saved at index 0 in a file. The likely reason is that when an element
    /// (e.g., a cell) in a XML doesn't explicitly define index of a format, the default value
    /// is 0 = this format.
    /// </summary>
    internal XLCellFormatValue DefaultCellFormat => _cellFormats[0];

    internal XLWorkbookStyles()
    {
        _numberFormats = new BiDictionary<int, XLNumberFormat>();
        _fontFormats = new BiDictionary<int, XLFontFormatValue>();
        _fillFormats = new BiDictionary<int, XLFillFormatValue>();
        _borderFormats = new BiDictionary<int, XLBorderFormatValue>();
        _alignmentFormats = new BiDictionary<int, XLAlignmentFormatValue>();
        _protectionFormats = new BiDictionary<int, XLProtectionFormatValue>();
        _cellFormats = new BiDictionary<int, XLCellFormatValue>();
        _cellStyles = new BiDictionary<StyleId, XLCellStyleValue>();
        _differentialFormats = new BiDictionary<int, XLDxfValue>();
        _tableStyles = new Dictionary<string, XLTableTheme>(XLHelper.NameComparer);
        _pivotStyles = new Dictionary<string, XLPivotTableStyle>(XLHelper.NameComparer);
    }

    internal IReadOnlyBiDictionary<int, XLNumberFormat> NumberFormats => _numberFormats;

    internal IReadOnlyBiDictionary<int, XLFontFormatValue> Fonts => _fontFormats;

    internal IReadOnlyBiDictionary<int, XLFillFormatValue> Fills => _fillFormats;

    internal IReadOnlyBiDictionary<int, XLBorderFormatValue> Borders => _borderFormats;

    internal IReadOnlyBiDictionary<int, XLCellFormatValue> CellFormats => _cellFormats;

    internal IReadOnlyBiDictionary<StyleId, XLCellStyleValue> CellStyles => _cellStyles;

    internal IReadOnlyBiDictionary<int, XLDxfValue> DifferentialFormats => _differentialFormats;

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
        NumberFormat = XLPredefinedFormat.FormatCodes[XLPredefinedFormat.General],
        Alignment = new XLAlignmentFormatValue()
        {
            Horizontal = XLAlignmentHorizontalValues.General,
            Vertical = XLAlignmentVerticalValues.Bottom,
            TextRotation = TextRotation.None,
            WrapText = false,
            Indent = 0,
            RelativeIndent = 0,
            JustifyLastLine = false,
            ShrinkToFit = false,
            ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent
        },
        Protection = new XLProtectionFormatValue
        {
            Locked = true,
            Hidden = false,
        },
        Fill = XLFillFormatValue.None,
        Border = XLBorderFormatValue.None,
        CellStyleId = null,
        IncludeQuotePrefix = false,
        PivotButton = false,
        CustomFormat = CellFormatComponents.None
    };

    internal void AddNumberFormat(int numFmtId, XLNumberFormat format)
    {
        _numberFormats.Add(numFmtId, format);
    }

    internal void AddUserDefinedNumberFormat(XLNumberFormat numberFormat)
    {
        var numFmtId = FirstUserDefinedNumberFormatIndex;
        if (_numberFormats.Count > 0)
            numFmtId = Math.Max(_numberFormats.Keys.Max() + 1, numFmtId);

        _numberFormats.Add(numFmtId, numberFormat);
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

    internal XLNumberFormat RegisterNumberFormat(XLNumberFormat numberFormat)
    {
        if (_numberFormats.TryGetValue(numberFormat, out var existingFormat))
            return existingFormat;

        AddUserDefinedNumberFormat(numberFormat);
        return numberFormat;
    }

    private XLAlignmentFormatValue GetRegisteredAlignmentFormat(XLAlignmentFormatValue original, Func<XLAlignmentFormatValue, XLAlignmentFormatValue> modify)
    {
        return RegisterAlignmentFormat(modify(original));
    }

    internal XLAlignmentFormatValue RegisterAlignmentFormat(XLAlignmentFormatValue alignment)
    {
        if (_alignmentFormats.TryGetValue(alignment, out var existingAlignment))
            return existingAlignment;

        _alignmentFormats.Add(_alignmentFormats.Count, alignment);
        return alignment;
    }

    internal XLProtectionFormatValue RegisterProtectionFormat(XLProtectionFormatValue protection)
    {
        if (_protectionFormats.TryGetValue(protection, out var existingProtection))
            return existingProtection;

        _protectionFormats.Add(_protectionFormats.Count, protection);
        return protection;
    }

    /// <summary>
    /// Get a font format that is stored in the internal structures of the styles class. The font
    /// format is created by modification of existing font format. This is essential for saving,
    /// all formats must be registered in the styles class. 
    /// </summary>
    internal XLFontFormatValue GetRegisteredFontFormat(XLFontFormatValue original, Func<XLFontFormatValue, XLFontFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterFontFormat(modified);
    }

    internal XLFontFormatValue RegisterFontFormat(XLFontFormatValue font)
    {
        if (_fontFormats.TryGetValue(font, out var existingFont))
            return existingFont;

        AddFontFormat(font);
        return font;
    }

    internal XLCellFormatValue GetModifiedFormat(XLCellFormatValue originalFormat, XLNumberFormat numberFormat)
    {
        var modifiedNumberFormat = RegisterNumberFormat(numberFormat);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { NumberFormat = modifiedNumberFormat });
        return modifiedFormat;
    }

    internal XLCellFormatValue GetModifiedFormat(XLCellFormatValue originalFormat, Func<XLAlignmentFormatValue, XLAlignmentFormatValue> modify)
    {
        var modifiedAlignment = GetRegisteredAlignmentFormat(originalFormat.Alignment, modify);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { Alignment = modifiedAlignment });
        return modifiedFormat;
    }

    internal XLCellFormatValue GetModifiedFormat(XLCellFormatValue originalFormat, Func<XLFontFormatValue, XLFontFormatValue> modify)
    {
        var modifiedFont = GetRegisteredFontFormat(originalFormat.Font, modify);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { Font = modifiedFont });
        return modifiedFormat;
    }

    internal XLCellFormatValue GetModifiedFormat(XLCellFormatValue originalFormat, Func<XLBorderFormatValue, XLBorderFormatValue> modify)
    {
        var modifiedBorder = GetRegisteredBorderFormat(originalFormat.Border, modify);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { Border = modifiedBorder });
        return modifiedFormat;
    }

    internal XLFillFormatValue GetRegisteredFillFormat(XLFillFormatValue original, Func<XLFillFormatValue, XLFillFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterFillFormat(modified);
    }

    private XLFillFormatValue RegisterFillFormat(XLFillFormatValue fill)
    {
        if (_fillFormats.TryGetValue(fill, out var existingFill))
            return existingFill;

        AddFillFormat(fill);
        return fill;
    }

    internal XLBorderFormatValue GetRegisteredBorderFormat(XLBorderFormatValue original, Func<XLBorderFormatValue, XLBorderFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterBorderFormat(modified);
    }

    private XLBorderFormatValue RegisterBorderFormat(XLBorderFormatValue border)
    {
        if (_borderFormats.TryGetValue(border, out var existingBorder))
            return existingBorder;

        AddBorderFormat(border);
        return border;
    }

    internal XLCellFormatValue GetRegisteredCellFormat(XLCellFormatValue original, Func<XLCellFormatValue, XLCellFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterCellFormat(modified);
    }

    internal XLCellFormatValue RegisterCellFormat(XLCellFormatValue cellFormat)
    {
        if (_cellFormats.TryGetValue(cellFormat, out var existing))
            return existing;

        Debug.Assert(_numberFormats.ContainsValue(cellFormat.NumberFormat));
        Debug.Assert(_alignmentFormats.TryGetValue(cellFormat.Alignment, out var registeredAlignment) && ReferenceEquals(cellFormat.Alignment, registeredAlignment));
        Debug.Assert(_protectionFormats.TryGetValue(cellFormat.Protection, out var registeredProtection) && ReferenceEquals(cellFormat.Protection, registeredProtection));
        Debug.Assert(_fontFormats.TryGetValue(cellFormat.Font, out var registeredFont) && ReferenceEquals(cellFormat.Font, registeredFont));
        Debug.Assert(_fillFormats.TryGetValue(cellFormat.Fill, out var registeredFill) && ReferenceEquals(cellFormat.Fill, registeredFill));
        Debug.Assert(_borderFormats.TryGetValue(cellFormat.Border, out var registeredBorder) && ReferenceEquals(cellFormat.Border, registeredBorder));
        AddFormat(cellFormat);
        return cellFormat;
    }

    /// <summary>
    /// Get registered format equal to <paramref name="format"/> from the styles. Generally for copying formats other workbooks.
    /// </summary>
    internal XLCellFormatValue GetRegisteredCellFormat(XLCellFormatValue format)
    {
        // If format is already registered, all its components must be registered too.
        if (_cellFormats.TryGetValue(format, out var existing))
            return existing;

        // TODO: If format is from different workbook, we should copy style if from different workbook and this workbook doesn't already have a style with same name
        StyleId? formatStyleId  = format.CellStyleId is { } cellStyleId && _cellStyles.ContainsKey(cellStyleId) ? cellStyleId : null;

        // We have to create new one, because some components may already exist here
        var cellFormat = new XLCellFormatValue
        {
            NumberFormat = RegisterNumberFormat(format.NumberFormat),
            Alignment = RegisterAlignmentFormat(format.Alignment),
            Protection = RegisterProtectionFormat(format.Protection),
            Font = RegisterFontFormat(format.Font),
            Fill = RegisterFillFormat(format.Fill),
            Border = RegisterBorderFormat(format.Border),
            CellStyleId = formatStyleId,
            IncludeQuotePrefix = format.IncludeQuotePrefix,
            PivotButton = format.PivotButton,
            CustomFormat = format.CustomFormat
        };

        AddFormat(cellFormat);
        return cellFormat;
    }

    /// <summary>
    /// Get a differential format that is stored in the internal structures of the styles class.
    /// The differential format is created by modification of existing dxf format. This is
    /// essential for saving, all formats must be registered in the styles class. 
    /// </summary>
    internal XLDxfValue GetRegisteredDxFormat(XLDxfValue original, Func<XLDxfValue, XLDxfValue> modify)
    {
        var modified = modify(original);
        return RegisterDxFormat(modified);
    }

    /// <summary>
    /// Register dxf from potentially different workbook into this workbook.
    /// </summary>
    /// <returns>Registered instance.</returns>
    internal XLDxfValue RegisterDxFormat(XLDxfValue dxf)
    {
        if (_differentialFormats.TryGetValue(dxf, out var existingDxf))
            return existingDxf;

        AddDifferentialFormat(dxf);
        return dxf;
    }

    /// <summary>
    /// Initialize styles component so it does contain styles suitable for a new workbook.
    /// </summary>
    internal void Initialize()
    {
        DefaultTableStyle = XLTableTheme.TableStyleMedium2.ToString();
        DefaultPivotStyle = nameof(XLPivotTableTheme.PivotStyleLight16);

        foreach (var (numFmtId, formatCode) in XLPredefinedFormat.FormatCodes)
            AddNumberFormat(numFmtId, formatCode);

        var normalStyle = DefaultNormalStyle;
        AddFontFormat(normalStyle.Font);
        AddFillFormat(XLFillFormatValue.None);
        AddFillFormat(XLFillFormatValue.Gray125);
        AddBorderFormat(XLBorderFormatValue.None);
        AddCellStyle(0, normalStyle);

        var defaultFormat = XLCellFormatValue.FromStyle(0, normalStyle);
        DefaultFormat = GetRegisteredCellFormat(defaultFormat);
    }
}
