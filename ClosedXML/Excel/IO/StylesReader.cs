using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel.Formatting;
using ClosedXML.IO;
using ClosedXML.Utils;
using TS = ClosedXML.Excel.Formatting.XLTableStyleRegionValues;
using PTS = ClosedXML.Excel.Formatting.XLPivotStyleRegionValues;

namespace ClosedXML.Excel.IO;

internal partial class StylesReader
{
    private readonly XmlTreeReader _reader;
    private readonly XLWorkbookStyles _styles;
    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly SequentialNameGenerator _styleNameGenerator = new("Style ", 1);

    // Currently read CT_TableStyle element
    private Dictionary<TS, (XLDxfValue Dxf, int BandSize)> _currentTableStyle = new();
    private Dictionary<PTS, (XLDxfValue Dxf, int BandSize)> _currentPivotStyle = new();

    /// <summary>
    /// Style formats from <c>cellStyleXfs</c>.
    /// </summary>
    private List<XLCellFormatValue> _styleFormats = new();

    public StylesReader(XmlTreeReader reader, XLWorkbookStyles styles)
    {
        _reader = reader;
        _styles = styles;
    }

    internal void Load()
    {
        _reader.Open("styleSheet", _ns);
        ParseStylesheet("styleSheet");

        LoadDefaultFormat();
    }

    private void LoadDefaultFormat()
    {
        // Normal style is technically optional.
        var normalStyle = _styles.CellStyles.Select(x => x.Value).SingleOrDefault(x => x.BuiltInStyle == BuiltInStyleValues.Normal);
        if (normalStyle is null)
            return;

        // Default format has mostly unchangeable values regardless of workbook styles, but it
        // does adjust font name and size from normal style on load.
        var defaultFontName = normalStyle.Font?.Name ?? _styles.DefaultFormat.Font?.Name;
        var defaultFontSize = normalStyle.Font?.Size ?? _styles.DefaultFormat.Font?.Size;
        _styles.DefaultFormat = _styles.DefaultFormat with
        {
            Font = _styles.DefaultFormat!.Font! with
            {
                Name = defaultFontName,
                Size = defaultFontSize
            }
        };
    }

    private void ParseStylesheet(string elementName)
    {
        if (_reader.TryOpen("numFmts", _ns))
        {
            ParseNumFmts("numFmts");
        }

        // The spec says that the predefined formats have "formatCode value [..] implied rather
        // than explicitly saved in the file."... so if there was something saved, it should have
        // be used. If something was saved, it's explicit and explicit (generally) has preference
        // over implicit. It needs to be added after numFmts, but before cellStyleXfs/cellXfs.
        AddImpliedNumberFormats();

        if (_reader.TryOpen("fonts", _ns))
        {
            ParseFonts("fonts");
        }
        if (_reader.TryOpen("fills", _ns))
        {
            ParseFills("fills");
        }
        if (_reader.TryOpen("borders", _ns))
        {
            ParseBorders("borders");
        }
        if (_reader.TryOpen("cellStyleXfs", _ns))
        {
            ParseCellStyleXfs("cellStyleXfs");
        }

        var cellFormats = new List<(XLCellFormatValue Format, int? CellStyleXfId)>();
        if (_reader.TryOpen("cellXfs", _ns))
        {
            cellFormats = ParseCellXfs("cellXfs");
        }

        var cellStyles = new Dictionary<int, XLCellStyleValue>();
        if (_reader.TryOpen("cellStyles", _ns))
        {
            cellStyles = ParseCellStyles("cellStyles");
        }

        RepairMissingStyles(cellStyles);
        AddCellStyles(cellStyles);
        AddFormats(cellFormats, cellStyles);

        if (_reader.TryOpen("dxfs", _ns))
        {
            ParseDxfs("dxfs");
        }
        if (_reader.TryOpen("tableStyles", _ns))
        {
            ParseTableStyles("tableStyles");
        }
        if (_reader.TryOpen("colors", _ns))
        {
            ParseColors("colors");
        }
        if (_reader.TryOpen("extLst", _ns))
        {
            ParseExtensionList("extLst");
        }
        _reader.Close(elementName, _ns);
    }

    private void AddImpliedNumberFormats()
    {
        // Add predefined formats, so we can treat predefined number formats and
        // user defined number formats the same way.
        foreach (var (numFmtId, formatCode) in XLPredefinedFormat.FormatCodes)
        {
            if (!_styles.NumberFormats.ContainsKey(numFmtId))
            {
                _styles.AddNumberFormat(numFmtId, formatCode);
            }
        }
    }

    private void RepairMissingStyles(Dictionary<int, XLCellStyleValue> cellStyles)
    {
        // Because cellStyleXfs might be referenced from cell formats, each one must be converted
        // to a cell style. If the cellStyles didn't contain a record for any cellStyleXf, add it.
        for (var cellStyleXfId = 0; cellStyleXfId < _styleFormats.Count; ++cellStyleXfId)
        {
            if (!cellStyles.ContainsKey(cellStyleXfId))
            {
                var format = _styleFormats[cellStyleXfId];
                var generatedName = _styleNameGenerator.NextUnusedStyleName();
                cellStyles.Add(cellStyleXfId, new XLCellStyleValue
                {
                    Name = generatedName,
                    BuiltInStyle = null,
                    Hidden = false,
                    NumberFormat = format.NumberFormat,
                    Alignment = format.Alignment,
                    Protection = format.Protection,
                    Font = format.Font,
                    Fill = format.Fill,
                    Border = format.Border,
                    IncludedComponents = CellFormatComponents.All
                });
            }
        }
    }

    private void AddCellStyles(Dictionary<int, XLCellStyleValue> cellStyles)
    {
        foreach (var (cellStyleXfId, cellStyle) in cellStyles)
            _styles.AddCellStyle(cellStyleXfId, cellStyle);
    }

    private void AddFormats(List<(XLCellFormatValue Format, int? CellStyleXfId)> cellFormats, Dictionary<int, XLCellStyleValue> cellStyles)
    {
        // At the time when cellXf were parsed, cell styles weren't resolved. Resolve them now.
        for (var xfId = 0; xfId < cellFormats.Count; ++xfId)
        {
            if (cellFormats[xfId].CellStyleXfId is { } cellStyleXfId)
            {
                // Make sure a the styleId is valid. The styleId for cell formats is read from file
                // and thus could be invalid. We don't want to crash later.
                if (!cellStyles.ContainsKey(cellStyleXfId))
                    throw PartStructureException.InvalidAttributeValue();

                var cellFormat = cellFormats[xfId].Format with { CellStyleId = cellStyleXfId };
                cellFormats[xfId] = (cellFormat, cellStyleXfId);
            }
        }

        foreach (var (cellFormat, _) in cellFormats)
            _styles.AddFormat(cellFormat);
    }

    private (int NumFmtId, string FormatCode) OnNumFmtParsed(uint numFmtId, string formatCode)
    {
        return (checked((int)numFmtId), formatCode);
    }

    partial void OnNumFmtsParsed(List<(int NumFmtId, string FormatCode)> numFmt, uint? count)
    {
        foreach (var (numFmtId, formatCode) in numFmt)
        {
            // Even if numFmtId is predefined, store the supplied value. Excel accepts predefined
            // numFmtId and uses supplied format instead of predefined format. It fixes situation
            // during save, where such items are saved to user supplied numFmtId range.
            _styles.AddNumberFormat(numFmtId, formatCode);
        }
    }

    private XLFontFormatValue ParseFont(string elementName)
    {
        // Font is mostly buggy specification. Excel basically chokes on anything but a sequence,
        // but standard requires an unbound choice where elements can repeat.
        XLFontName? fontName = null;
        XLFontCharSet? fontCharset = null;
        XLFontFamilyNumberingValues? fontFamily = null;
        bool? fontBold = null, fontItalic = null, fontStrikethrough = null, fontOutline = null, fontShadow = null, fontCondense = null, fontExtend = null;
        XLColor? fontColor = null;
        XLFontSize? fontSize = null;
        XLFontUnderlineValues? fontUnderline = null;
        XLFontVerticalTextAlignmentValues? fontVerticalAlignment = null;
        XLFontScheme? fontScheme = null;
        while (!_reader.TryClose(elementName, _ns))
        {
            if (_reader.TryReadXStringValElement("name", _ns, out var name))
            {
                fontName = name;
            }
            else if (_reader.TryReadIntValElement("charset", _ns, out var charset))
            {
                fontCharset = (XLFontCharSet)charset;
            }
            else if (_reader.TryReadIntValElement("family", _ns, out var family))
            {
                // Bug in the spec. Spec says that it has values 0-14 and doesn't specify meaning
                // for the numerical values. It's supposed to refer to the same enum ST_FontFamily
                // as in WordML. The OI-29500 fixes this problem:
                // "Excel restricts the value of this attribute to be at least 0 and at most 5."
                fontFamily = family switch
                {
                    >= 0 and <= 5 => (XLFontFamilyNumberingValues)family,
                    > 5 and <= 14 => XLFontFamilyNumberingValues.NotApplicable,
                    _ => throw PartStructureException.InvalidAttributeFormat(),
                };
            }
            else if (_reader.TryReadBoolElement("b", _ns, out var b))
            {
                fontBold = b;
            }
            else if (_reader.TryReadBoolElement("i", _ns, out var i))
            {
                fontItalic = i;
            }
            else if (_reader.TryReadBoolElement("strike", _ns, out var strike))
            {
                fontStrikethrough = strike;
            }
            else if (_reader.TryReadBoolElement("outline", _ns, out var outline))
            {
                fontOutline = outline;
            }
            else if (_reader.TryReadBoolElement("shadow", _ns, out var shadow))
            {
                fontShadow = shadow;
            }
            else if (_reader.TryReadBoolElement("condense", _ns, out var condense))
            {
                fontCondense = condense;
            }
            else if (_reader.TryReadBoolElement("extend", _ns, out var extend))
            {
                fontExtend = extend;
            }
            else if (_reader.TryReadColor("color", _ns, out var color))
            {
                fontColor = color;
            }
            else if (_reader.TryOpen("sz", _ns))
            {
                var fontSizePt = _reader.GetDouble("val");
                _reader.Close("sz", _ns);
                fontSize = XLFontSize.FromPoints(fontSizePt);
            }
            else if (_reader.TryOpen("u", _ns))
            {
                var underline = _reader.GetOptionalEnum<XLFontUnderlineValues>("val") ?? XLFontUnderlineValues.Single;
                _reader.Close("u", _ns);
                fontUnderline = underline;
            }
            else if (_reader.TryReadEnumValElement<XLFontVerticalTextAlignmentValues>("vertAlign", _ns, out var vertAlign))
            {
                fontVerticalAlignment = vertAlign;
            }
            else if (_reader.TryReadEnumValElement<XLFontScheme>("scheme", _ns, out var scheme))
            {
                fontScheme = scheme;
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(_reader);
            }
        }

        var fontFormat = new XLFontFormatValue
        {
            Name = fontName,
            Charset = fontCharset,
            Family = fontFamily,
            Bold = fontBold,
            Italic = fontItalic,
            Strikethrough = fontStrikethrough,
            Outline = fontOutline,
            Shadow = fontShadow,
            Condense = fontCondense,
            Extend = fontExtend,
            Color = fontColor,
            Size = fontSize,
            Underline = fontUnderline,
            VerticalAlignment = fontVerticalAlignment,
            Scheme = fontScheme,
        };
        _styles.AddFontFormat(fontFormat);
        return fontFormat;
    }

    private XLFillFormatValue OnFillParsed(XLFillFormatValue? patternFill, XLFillFormatValue? gradientFill)
    {
        var fillFormat = patternFill ?? gradientFill ?? XLFillFormatValue.Empty;
        _styles.AddFillFormat(fillFormat);
        return fillFormat;
    }

    private XLFillFormatValue OnPatternFillParsed(XLColor? fgColor, XLColor? bgColor, XLFillPatternValues? patternType)
    {
        // There is a discrepancy between <fill> interpretation for a solid fill:
        // * cell fill: Pattern color is the one used for fill, the background color is ignored
        // * dxf fill: Pattern color is ignored, the background color is used for fill
        // The GUI in both cases says that the background color is the one that is used. Therefore
        // use background is correct per GUI. The problem is that ClosedXML historically says
        // the pattern color is the one that is used. This sucks, I have to live with it.
        // The alternative is a breaking change with no benefit.
        //
        // The other difference between dxf and cell fill is the default pattern type. Spec and
        // OI-29500 is silent, but Excel uses solid fill for dxf and none for cell fill.
        if (_reader.Context[^2] == "dxf")
        {
            var pattern = patternType ?? XLFillPatternValues.Solid;
            if (pattern == XLFillPatternValues.Solid)
            {
                // Fix solid pattern discrepancy for dxf
                var solidFill = new XLPatternFill
                {
                    PatternColor = bgColor ?? XLColor.NoColor,
                    BackgroundColor = fgColor ?? XLColor.NoColor,
                    PatternType = XLFillPatternValues.Solid,
                };
                return new XLFillFormatValue(solidFill);
            }

            var patternFill = new XLPatternFill
            {
                PatternColor = fgColor ?? XLColor.NoColor,
                BackgroundColor = bgColor ?? XLColor.NoColor,
                PatternType = pattern,
            };
            return new XLFillFormatValue(patternFill);
        }
        else
        {
            // Pattern for cell style fill
            var patternFill = new XLPatternFill
            {
                PatternColor = fgColor ?? XLColor.NoColor,
                BackgroundColor = bgColor ?? XLColor.NoColor,
                PatternType = patternType ?? XLFillPatternValues.None,
            };
            return new XLFillFormatValue(patternFill);
        }
    }

    private XLFillFormatValue OnGradientFillParsed(List<(FractionOfOne Value, XLColor Color)> stop, XLGradientType type, double degree, double left, double right, double top, double bottom)
    {
        var stops = stop.ToDictionary(x => x.Value, x => x.Color);
        switch (type)
        {
            case XLGradientType.Linear:
                return new XLFillFormatValue(new XLLinearGradientFill
                {
                    Stops = stops,
                    Degrees = degree,
                });
            case XLGradientType.Path:
                return new XLFillFormatValue(new XLPathGradientFill
                {
                    Stops = stops,
                    InnerLeft = left,
                    InnerRight = right,
                    InnerTop = top,
                    InnerBottom = bottom
                });
            default:
                throw new UnreachableException();
        }
    }

    private (FractionOfOne Position, XLColor Color) OnGradientStopParsed(XLColor color, double position)
    {
        // Spec requires stop positions to be 0..1, but doesn't have a type for that. Excel repairs workbook when it receives values outside 0..1.
        return (position, color);
    }

    private XLBorderFormatValue OnBorderParsed(XLBorderLine? left, XLBorderLine? right, XLBorderLine? top, XLBorderLine? bottom, XLBorderLine? diagonal, XLBorderLine? vertical, XLBorderLine? horizontal, bool? diagonalUp, bool? diagonalDown, bool outline)
    {
        var borderFormat = new XLBorderFormatValue
        {
            Left = left,
            Right = right,
            Top = top,
            Bottom = bottom,
            Diagonal = diagonal,
            Vertical = vertical,
            Horizontal = horizontal,
            DiagonalUp = diagonalUp ?? false, // OI-29500: Excel uses false as default value
            DiagonalDown = diagonalDown ?? false, // OI-29500: Excel uses false as default value
            Outline = outline
        };
        _styles.AddBorderFormat(borderFormat);
        return borderFormat;
    }

    private XLBorderLine OnBorderPrParsed(XLColor? color, XLBorderStyleValues style)
    {
        return new XLBorderLine(color ?? XLColor.NoColor, style);
    }

    partial void OnCellStyleXfsParsed(List<(XLCellFormatValue Format, int? CellStyleXfId)> xf, uint? count)
    {
        _styleFormats = xf.Select(x => x.Format).ToList();
    }

    private (XLCellFormatValue Format, int? CellStyleXfId) OnXfParsed(XLAlignmentFormatValue? alignment, XLProtectionFormatValue? protection, uint? numFmtId, uint? fontId, uint? fillId, uint? borderId, uint? xfId, bool quotePrefix, bool pivotButton, bool? applyNumberFormat, bool? applyFont, bool? applyFill, bool? applyBorder, bool? applyAlignment, bool? applyProtection)
    {
        // When xf is parsed, all number formats, fonts, fills and borders should already be read.
        string? numberFormat = null;
        if (numFmtId is not null)
            numberFormat = _styles.NumberFormats.GetValueOrDefault(checked((int)numFmtId));

        var font = fontId is not null ? _styles.Fonts[checked((int)fontId)] : null;
        var fill = fillId is not null ? _styles.Fills[checked((int)fillId)] : null;
        var border = borderId is not null ? _styles.Borders[checked((int)borderId)] : null;

        // Excel doesn't actually use the apply* for xf, but at least it writes as if it did. It
        // actually checks whether the id is same for xf and a style and if it is, the aspect
        // should be from a style. Excel is doesn't use it, other producers might.

        // Cell format has default apply* false (interpreted as "does format has its own custom format for this aspect")
        // Style has default apply* true (interpreted as "does style define this aspect")
        // The apply* attributes have default value true for cellStyleXfs and false for cellXfs.
        var isStyleXf = _reader.Context[^1] == "cellStyleXfs";
        var isCellFormat = !isStyleXf;
        var defaultApply = !isCellFormat;
        var components = CellFormatComponents.None;

        if (applyNumberFormat ?? defaultApply)
            components |= CellFormatComponents.NumberFormat;

        if (applyFont ?? defaultApply)
            components |= CellFormatComponents.Font;

        if (applyFill ?? defaultApply)
            components |= CellFormatComponents.Fill;

        if (applyBorder ?? defaultApply)
            components |= CellFormatComponents.Border;

        if (applyAlignment ?? defaultApply)
            components |= CellFormatComponents.Alignment;

        if (applyProtection ?? defaultApply)
            components |= CellFormatComponents.Protection;

        var format = new XLCellFormatValue
        {
            NumberFormat = numberFormat,
            Alignment = alignment,
            Protection = protection,
            Font = font,
            Fill = fill,
            Border = border,
            CellStyleId = null, // The style is set once cell styles are resolved
            IncludeQuotePrefix = quotePrefix,
            PivotButton = pivotButton,
            CustomFormat = components
        };
        return (format, checked((int?)xfId));
    }

    private List<(XLCellFormatValue Format, int? CellStyleXfId)> OnCellXfsParsed(List<(XLCellFormatValue Format, int? CellStyleXfId)> xf, uint? count)
    {
        return xf;
    }

    private (int XfId, XLCellStyleValue Style) OnCellStyleParsed(string? name, uint xfId, uint? builtinId, uint? iLevel, bool? hidden, bool? customBuiltin)
    {
        // The quotePrefix and pivotButton attributes of the cellStyleXf are not applied, plus
        // the xfId of the cellStyle is ignored. The OI-29500 also requires uniqueness of xfId
        // in the cellStyle elements, although Excel can load such workbook and has several
        // "linked" styles. It's likely the separation into two elements is based on the internal
        // structure inside the Excel.

        // Fill dummy name for the style
        if (string.IsNullOrWhiteSpace(name))
        {
            name = _styleNameGenerator.NextUnusedStyleName();
        }
        else
        {
            _styleNameGenerator.AddName(name);
        }

        var cellStyleFormat = _styleFormats[checked((int)xfId)];

        // If the built in style is an outline style, expand it to avoid ugly representation.
        // The spec has only one only builtIn id for all RowLevel* styles (1) and one builtIn
        // id for all ColLevel* styles (2). Since the iLevel/outlineLevel is used only for
        // the RowLevel/ColLevel (OI-29500), expand the builtIn+iLevel for Row/Col level into
        // a separate builtIn styles (101-107 for RowLevel1-7, 201-207 for ColumnLevel1-7).
        if (builtinId is 1 or 2)
            builtinId = builtinId.Value * 100 + 1 + iLevel ?? 0;

        // BuiltIn must be among defined built-in styles ("implementers should restrict the content
        // of this attribute to enumerations present in the list")
        if (builtinId is not null && !Enum.IsDefined(typeof(BuiltInStyleValues), checked((int)builtinId.Value)))
        {
            throw PartStructureException.InvalidAttributeFormat();
        }

        // The apply* attributes have default `true` for cellStyleXfs and `false` for cellXfs.
        // We already took care of correct default value during the parsing of <xf>, so we don't
        // have to deal with it here.
        var styleIncludesComponents = cellStyleFormat.CustomFormat;
        var cellStyle = new XLCellStyleValue
        {
            Name = name,
            BuiltInStyle = builtinId is not null ? (BuiltInStyleValues)builtinId.Value : null,
            Hidden = hidden ?? false,
            NumberFormat = cellStyleFormat.NumberFormat,
            Alignment = cellStyleFormat.Alignment,
            Protection = cellStyleFormat.Protection,
            Font = cellStyleFormat.Font,
            Fill = cellStyleFormat.Fill,
            Border = cellStyleFormat.Border,
            IncludedComponents = styleIncludesComponents
        };

        return (checked((int)xfId), cellStyle);
    }

    private Dictionary<int, XLCellStyleValue> OnCellStylesParsed(List<(int CellStyleXfId, XLCellStyleValue Style)> cellStyle, uint? count)
    {
        var cellStyles = new Dictionary<int, XLCellStyleValue>();
        foreach (var (cellStyleXfId, style) in cellStyle)
        {
            // Multiple cell styles use same style formatting - split them, so each one uses
            // separate formatting. I considered removing duplicates, but it could mean that I
            // also might remove normal style, which is not desirable.
            if (!cellStyles.ContainsKey(cellStyleXfId))
            {
                cellStyles.Add(cellStyleXfId, style);
            }
            else
            {
                _styleFormats.Add(_styleFormats[cellStyleXfId]);
                var newCellStyleXfId = _styleFormats.Count - 1;
                cellStyles.Add(newCellStyleXfId, style);
            }
        }

        return cellStyles;
    }

    private XLAlignmentFormatValue OnCellAlignmentParsed(XLAlignmentHorizontalValues? horizontal, XLAlignmentVerticalValues vertical, uint? textRotation, bool? wrapText, uint? indent, int? relativeIndent, bool? justifyLastLine, bool? shrinkToFit, uint? readingOrder)
    {
        if (readingOrder is not null && readingOrder is not (0 or 1 or 2))
            throw PartStructureException.InvalidAttributeFormat();

        var normalizedTextRotation = OpenXmlHelper.NormalizeRotation(textRotation ?? 0);
        return new XLAlignmentFormatValue
        {
            Horizontal = horizontal ?? XLAlignmentHorizontalValues.General,
            Vertical = vertical,
            TextRotation = new TextRotation(normalizedTextRotation),
            WrapText = wrapText ?? false,
            Indent = indent is not null ? checked((int)indent.Value) : 0,
            RelativeIndent = relativeIndent ?? 0,
            JustifyLastLine = justifyLastLine ?? false,
            ShrinkToFit = shrinkToFit ?? false,
            ReadingOrder = readingOrder is not null ? (XLAlignmentReadingOrderValues)readingOrder.Value : XLAlignmentReadingOrderValues.ContextDependent
        };
    }

    partial void OnDxfParsed(XLFontFormatValue? font, (int NumFmtId, string FormatCode)? numFmt, XLFillFormatValue? fill, XLAlignmentFormatValue? alignment, XLBorderFormatValue? border, XLProtectionFormatValue? protection)
    {
        var dxf = new XLDxfValue
        {
            NumberFormat = numFmt?.FormatCode,
            Font = font,
            Fill = fill,
            Alignment = alignment,
            Border = border,
            Protection = protection,
        };
        _styles.AddDifferentialFormat(dxf);
    }

    partial void OnTableStyleElementParsed((TS?, PTS?) type, uint size, uint? dxfId)
    {
        // Skip definition without a differential format
        if (dxfId is null)
            return;

        // Excel permits only 0-9
        if (size > 9)
            throw PartStructureException.InvalidAttributeFormat(nameof(size), size.ToString(), _reader);

        // If there is a duplicate definition for a type, last one wins
        var dxf = _styles.DifferentialFormats[checked((int)dxfId.Value)];

        if (type.Item1 is { } tableStyleRegion)
            _currentTableStyle[tableStyleRegion] = (dxf, (int)size);

        if (type.Item2 is { } pivotStyleRegion)
            _currentPivotStyle[pivotStyleRegion] = (dxf, (int)size);
    }

    partial void OnTableStyleParsed(string name, bool pivot, bool table, uint? count)
    {
        // Because of tableStyle element duality, we are filling both styles and
        // only insert types that have set flag.
        if (table)
        {
            var tableStyle = new XLTableTheme(name);
            foreach (var (region, (dxf, bandSize)) in _currentTableStyle)
                tableStyle.SetRegionFormat(region, dxf, bandSize);

            _styles.AddTableStyle(tableStyle);
        }

        _currentTableStyle = new Dictionary<TS, (XLDxfValue Dxf, int BandSize)>();

        if (pivot)
        {
            var pivotStyle = new XLPivotTableStyle(name);
            foreach (var (region, (dxf, bandSize)) in _currentPivotStyle)
                pivotStyle.SetRegionFormat(region, dxf, bandSize);

            _styles.AddPivotStyle(pivotStyle);
        }

        _currentPivotStyle = new Dictionary<PTS, (XLDxfValue Dxf, int BandSize)>();
    }

    partial void OnTableStylesParsed(uint? count, string? defaultTableStyle, string? defaultPivotStyle)
    {
        if (!string.IsNullOrEmpty(defaultTableStyle))
            _styles.DefaultTableStyle = defaultTableStyle;

        if (!string.IsNullOrEmpty(defaultPivotStyle))
            _styles.DefaultPivotStyle = defaultPivotStyle;
    }

    private uint OnRgbColorParsed(uint? rgb)
    {
        // Despite the name, it's ARGB. If not specified, use black (Excel supplies 0x00000000, but
        // Excel plays very fast and loose with transparency).
        return rgb ?? 0xFF000000;
    }

    partial void OnIndexedColorsParsed(List<uint> rgbColor)
    {
        _styles.SetIndexedColors(rgbColor);
    }

    partial void OnMRUColorsParsed(List<XLColor> color)
    {
        _styles.SetMruColors(color);
    }

    private XLColor ParseColor(string elementName)
    {
        return _reader.ParseColor(elementName, _ns);
    }

    private void ParseExtensionList(string elementName)
    {
        _reader.Skip(elementName);
    }

    private XLProtectionFormatValue OnCellProtectionParsed(bool? locked, bool? hidden)
    {
        // Defaults are from OI-29500
        return new XLProtectionFormatValue
        {
            Locked = locked ?? true,
            Hidden = hidden ?? false
        };
    }

    /// <summary>
    /// A mapping of <c>ST_TableStyleType</c>. Custom enum mapping due to table/pivot duality.
    /// </summary>
    private static readonly Dictionary<string, (TS?, PTS?)> TableStyleTypeMap = new()
    {
        { "wholeTable", (TS.WholeTable, PTS.WholeTable) },
        { "headerRow", (TS.HeaderRow, PTS.HeaderRow) },
        { "totalRow", (TS.TotalRow, PTS.GrandTotalRow) },
        { "firstColumn", (TS.FirstColumn, PTS.FirstColumn) },
        { "lastColumn", (TS.LastColumn, PTS.GrandTotalColumn) },
        { "firstRowStripe", (TS.FirstRowStripe, PTS.FirstRowStripe) },
        { "secondRowStripe", (TS.SecondRowStripe, PTS.SecondRowStripe) },
        { "firstColumnStripe", (TS.FirstColumnStripe, PTS.FirstColumnStripe) },
        { "secondColumnStripe", (TS.SecondColumnStripe, PTS.SecondColumnStripe) },
        { "firstHeaderCell", (TS.FirstHeaderCell, PTS.FirstHeaderCell) },
        { "lastHeaderCell", (TS.LastHeaderCell, null) },
        { "firstTotalCell", ( TS.FirstTotalCell, null) },
        { "lastTotalCell", ( TS.LastTotalCell, null) },
        { "firstSubtotalColumn", (null, PTS.SubtotalColumn1) },
        { "secondSubtotalColumn", (null,PTS.SubtotalColumn2) },
        { "thirdSubtotalColumn", (null, PTS.SubtotalColumn3) },
        { "firstSubtotalRow", (null, PTS.SubtotalRow1) },
        { "secondSubtotalRow", (null, PTS.SubtotalRow2) },
        { "thirdSubtotalRow", (null, PTS.SubtotalRow3) },
        { "blankRow", (null, PTS.BlankRow) },
        { "firstColumnSubheading", (null, PTS.ColumnSubheading1) },
        { "secondColumnSubheading", (null, PTS.ColumnSubheading2) },
        { "thirdColumnSubheading", (null, PTS.ColumnSubheading3) },
        { "firstRowSubheading", (null, PTS.RowSubheading1) },
        { "secondRowSubheading", (null, PTS.RowSubheading2) },
        { "thirdRowSubheading", (null, PTS.RowSubheading3) },
        { "pageFieldLabels", (null, PTS.PageFieldLabels) },
        { "pageFieldValues", (null, PTS.PageFieldValues) },
    };
}
