using ClosedXML.Excel.Formatting;
using ClosedXML.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Utils;

namespace ClosedXML.Excel.IO;

internal partial class StylesReader
{
    private readonly XmlTreeReader _reader;
    private readonly XLWorkbookStyles _styles;
    private readonly string _ns = OpenXmlConst.Main2006SsNs;

    public StylesReader(XmlTreeReader reader, XLWorkbookStyles styles)
    {
        _reader = reader;
        _styles = styles;
    }

    internal void Load()
    {
        _reader.Open("styleSheet", _ns);
        ParseStylesheet("styleSheet");

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

    private (int NumFmtId, string FormatCode) OnNumFmtParsed(uint numFmtId, string formatCode)
    {
        return (checked((int)numFmtId), formatCode);
    }

    partial void OnNumFmtsParsed(List<(int NumFmtId, string FormatCode)> numFmt, uint? count)
    {
        foreach (var (numFmtId, formatCode) in numFmt)
        {
            // Excel skips empty format codes.
            if (string.IsNullOrEmpty(formatCode))
                continue;

            // Even if numFmtId is predefined, store the supplied value. Excel accepts predefined
            // numFmtId and uses supplied format instead of predefined format. It fixes situation
            // during save, where such items are saved to user supplied numFmtId range.
            _styles.AddNumberFormat(numFmtId, formatCode);
        }
    }

    private void ParseFont(string elementName)
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

        var format = new XLFontFormat
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
        _styles.AddFontFormat(format);
    }

    partial void OnFillParsed(XLFillFormat? patternFill, XLFillFormat? gradientFill)
    {
        _styles.AddFillFormat(patternFill ?? gradientFill ?? XLFillFormat.Empty);
    }

    private XLFillFormat OnPatternFillParsed(XLColor? fgColor, XLColor? bgColor, XLFillPatternValues? patternType)
    {
        var patternFill = new XLPatternFill
        {
            PatternColor = fgColor ?? XLColor.NoColor,
            BackgroundColor = bgColor ?? XLColor.NoColor,
            PatternType = patternType ?? XLFillPatternValues.None,
        };
        return new XLFillFormat(patternFill);
    }

    private XLFillFormat OnGradientFillParsed(List<(FractionOfOne Value, XLColor Color)> stop, XLGradientType type, double degree, double left, double right, double top, double bottom)
    {
        var stops = stop.ToDictionary(x => x.Value, x => x.Color);
        switch (type)
        {
            case XLGradientType.Linear:
                return new XLFillFormat(new XLLinearGradientFill
                {
                    Stops = stops,
                    Degrees = degree,
                });
            case XLGradientType.Path:
                return new XLFillFormat(new XLPathGradientFill
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

    partial void OnBorderParsed(XLBorderLine? left, XLBorderLine? right, XLBorderLine? top, XLBorderLine? bottom, XLBorderLine? diagonal, XLBorderLine? vertical, XLBorderLine? horizontal, bool? diagonalUp, bool? diagonalDown, bool outline)
    {
        var borderFormat = new XLBorderFormat
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
    }

    private XLBorderLine OnBorderPrParsed(XLColor? color, XLBorderStyleValues style)
    {
        return new XLBorderLine(color ?? XLColor.NoColor, style);
    }

    partial void OnXfParsed(XLAlignmentFormat? alignment, XLProtectionFormat? protection, uint? numFmtId, uint? fontId, uint? fillId, uint? borderId, uint? xfId, bool quotePrefix, bool pivotButton, bool? applyNumberFormat, bool? applyFont, bool? applyFill, bool? applyBorder, bool? applyAlignment, bool? applyProtection)
    {
        // When xf is parsed, all number formats, fonts, fills and borders should already be read.
        // Skip cell style xfs for now.
        if (_reader.Context[^1] == "cellXfs")
        {
            // We are in cellXfs
            var font = fontId is not null ? _styles.Fonts[checked((int)fontId)] : null;
            var fill = fillId is not null ? _styles.Fills[checked((int)fillId)] : null;
            var border = borderId is not null ? _styles.Borders[checked((int)borderId)] : null;

            _styles.AddFormat(alignment, protection, font, fill, border);
        }
    }

    private XLAlignmentFormat OnCellAlignmentParsed(XLAlignmentHorizontalValues? horizontal, XLAlignmentVerticalValues vertical, uint? textRotation, bool? wrapText, uint? indent, int? relativeIndent, bool? justifyLastLine, bool? shrinkToFit, uint? readingOrder)
    {
        if (readingOrder is not null && readingOrder is not (0 or 1 or 2))
            throw PartStructureException.InvalidAttributeFormat();

        var normalizedTextRotation = OpenXmlHelper.NormalizeRotation(textRotation ?? 0);
        return new XLAlignmentFormat
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

    private XLColor ParseColor(string elementName)
    {
        return _reader.ParseColor(elementName, _ns);
    }

    private void ParseExtensionList(string elementName)
    {
        _reader.Skip(elementName);
    }

    private XLProtectionFormat OnCellProtectionParsed(bool? locked, bool? hidden)
    {
        // Defaults are from OI-29500
        return new XLProtectionFormat
        {
            Locked = locked ?? true,
            Hidden = hidden ?? false
        };
    }
}
