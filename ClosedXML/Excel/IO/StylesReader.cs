using ClosedXML.Excel.Formatting;
using ClosedXML.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel.IO;

internal partial class StylesReader
{
    private readonly XmlTreeReader _reader;
    private readonly XLWorkbookStyles _styles;
    private readonly string _ns = OpenXmlConst.Main2006SsNs;

    /// <summary>
    /// A marker for <c>xf</c>> parsing. The <c>cellStyleXfs</c> and <c>cellXfs</c> both use same
    /// element and even same name. This flag is used on the hook to differentiate them.
    /// </summary>
    private bool _insideCellXfs = false;

    public StylesReader(XmlTreeReader reader, XLWorkbookStyles styles)
    {
        _reader = reader;
        _styles = styles;
    }

    internal void Load()
    {
        _reader.Open("styleSheet", _ns);
        ParseStylesheet("styleSheet");
    }

    private void ParseFont(string elementName)
    {
        // Font is mostly buggy specification. Excel basically chokes on anything but a sequence,
        // but standard requires an unbound choice where elements can repeat.
        XLFontFormat format = default;
        while (!_reader.TryClose(elementName, _ns))
        {
            if (_reader.TryReadXStringValElement("name", _ns, out var fontName))
            {
                format = format with { Name = fontName };
            }
            else if (_reader.TryReadIntValElement("charset", _ns, out var charset))
            {
                format = format with { Charset = (XLFontCharSet?)charset };
            }
            else if (_reader.TryReadIntValElement("family", _ns, out var family))
            {
                format = format with { Family = (XLFontFamilyNumberingValues)family };
            }
            else if (_reader.TryReadBoolElement("b", _ns, out var b))
            {
                format = format with { Bold = b };
            }
            else if (_reader.TryReadBoolElement("i", _ns, out var i))
            {
                format = format with { Italic = i };
            }
            else if (_reader.TryReadBoolElement("strike", _ns, out var strike))
            {
                format = format with { Strikethrough = strike };
            }
            else if (_reader.TryReadBoolElement("outline", _ns, out var outline))
            {
                format = format with { Outline = outline };
            }
            else if (_reader.TryReadBoolElement("shadow", _ns, out var shadow))
            {
                format = format with { Shadow = shadow };
            }
            else if (_reader.TryReadBoolElement("condense", _ns, out var condense))
            {
                format = format with { Condense = condense };
            }
            else if (_reader.TryReadBoolElement("extend", _ns, out var extend))
            {
                format = format with { Extend = extend };
            }
            else if (_reader.TryReadColor("color", _ns, out var color))
            {
                format = format with { Color = color };
            }
            else if (_reader.TryOpen("sz", _ns))
            {
                var fontSizePt = _reader.GetDouble("val");
                _reader.Close("sz", _ns);
                format = format with { Size = XLFontSize.FromPoints(fontSizePt) };
            }
            else if (_reader.TryOpen("u", _ns))
            {
                var underline = _reader.GetOptionalEnum<XLFontUnderlineValues>("val") ?? XLFontUnderlineValues.Single;
                _reader.Close("u", _ns);
                format = format with { Underline = underline };
            }
            else if (_reader.TryReadEnumValElement<XLFontVerticalTextAlignmentValues>("vertAlign", _ns, out var vertAlign))
            {
                format = format with { VerticalAlignment = vertAlign };
            }
            else if (_reader.TryReadEnumValElement<XLFontScheme>("scheme", _ns, out var scheme))
            {
                format = format with { Scheme = scheme };
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(_reader);
            }
        }

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

    private void ParseCellXfs(string elementName)
    {
        _insideCellXfs = true;
        _reader.Open("xf", _ns);
        do
        {
            ParseXf("xf");
        }
        while (_reader.TryOpen("xf", _ns));
        _reader.Close(elementName, _ns);
        _insideCellXfs = false;
    }

    partial void OnXfParsed(uint? numFmtId, uint? fontId, uint? fillId, uint? borderId, uint? xfId, bool quotePrefix, bool pivotButton, bool? applyNumberFormat, bool? applyFont, bool? applyFill, bool? applyBorder, bool? applyAlignment, bool? applyProtection)
    {
        // When xf is parsed, all number formats, fonts, fills and borders should already be read.
        // Skip cell style xfs for now.
        if (_insideCellXfs)
        {
            // We are in cellXfs
            _styles.AddFormat(fontId, fillId, borderId);
        }
    }

    private XLColor ParseColor(string elementName)
    {
        return _reader.ParseColor(elementName, _ns);
    }

    private void ParseExtensionList(string elementName)
    {
        _reader.Skip(elementName);
    }
}
