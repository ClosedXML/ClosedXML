#nullable enable

using System.Collections.Generic;
using ClosedXML.IO;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel.IO;

internal partial class StylesReader
{
    private Xpr ParseNumFmts(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        var numFmt = new List<(int NumFmtId, XLNumberFormat Format)>();
        while (ParseNumFmt("numFmt", _ns) is { IsSuccess: true} numFmtItem)
        {
            numFmt.Add(numFmtItem.Value);
        }
        _reader.Close(elementName, ns);

        OnNumFmtsParsed(numFmt, count);
        return Xpr.Success();
    }

    partial void OnNumFmtsParsed(List<(int NumFmtId, XLNumberFormat Format)> numFmt, uint? count);

    private Xpr<(int NumFmtId, XLNumberFormat Format)> ParseNumFmt(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(int NumFmtId, XLNumberFormat Format)>();
        }

        var numFmtId = _reader.GetUInt("numFmtId");
        var formatCode = _reader.GetXString("formatCode");

        _reader.Close(elementName, ns);

        return Xpr.From(OnNumFmtParsed(numFmtId, formatCode));
    }

    private Xpr ParseFonts(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        var font = new List<XLDifferentialFontValue>();
        while (ParseFont("font", _ns) is { IsSuccess: true} fontItem)
        {
            font.Add(fontItem.Value);
        }
        _reader.Close(elementName, ns);

        OnFontsParsed(font, count);
        return Xpr.Success();
    }

    partial void OnFontsParsed(List<XLDifferentialFontValue> font, uint? count);

    private Xpr ParseFills(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        var fill = new List<XLFillFormatValue>();
        while (ParseFill("fill", _ns) is { IsSuccess: true} fillItem)
        {
            fill.Add(fillItem.Value);
        }
        _reader.Close(elementName, ns);

        OnFillsParsed(fill, count);
        return Xpr.Success();
    }

    partial void OnFillsParsed(List<XLFillFormatValue> fill, uint? count);

    private Xpr<XLFillFormatValue> ParseFill(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFillFormatValue>();
        }

        XLFillFormatValue? choice;
        if (ParsePatternFill("patternFill", _ns) is { IsSuccess: true } patternFill)
        {
            choice = OnFillPatternFillParsed(patternFill.Value);
        }
        else if (ParseGradientFill("gradientFill", _ns) is { IsSuccess: true } gradientFill)
        {
            choice = OnFillGradientFillParsed(gradientFill.Value);
        }
        else
        {
            choice = default;
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnFillParsed(choice));
    }

    private Xpr<XLFillFormatValue> ParsePatternFill(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFillFormatValue>();
        }

        var patternType = _reader.GetOptionalEnum<XLFillPatternValues>("patternType");

        var fgColorResult = ParseColor("fgColor", _ns);
        var fgColor = fgColorResult.IsSuccess ? fgColorResult.Value : default(XLColor?);
        var bgColorResult = ParseColor("bgColor", _ns);
        var bgColor = bgColorResult.IsSuccess ? bgColorResult.Value : default(XLColor?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnPatternFillParsed(fgColor, bgColor, patternType));
    }

    private Xpr<XLFillFormatValue> ParseGradientFill(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFillFormatValue>();
        }

        var type = _reader.GetOptionalEnum<XLGradientType>("type") ?? XLGradientType.Linear;
        var degree = _reader.GetOptionalDouble("degree") ?? 0;
        var left = _reader.GetOptionalDouble("left") ?? 0;
        var right = _reader.GetOptionalDouble("right") ?? 0;
        var top = _reader.GetOptionalDouble("top") ?? 0;
        var bottom = _reader.GetOptionalDouble("bottom") ?? 0;

        var stop = new List<(FractionOfOne Value, XLColor Color)>();
        while (ParseGradientStop("stop", _ns) is { IsSuccess: true} stopItem)
        {
            stop.Add(stopItem.Value);
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnGradientFillParsed(stop, type, degree, left, right, top, bottom));
    }

    private Xpr<(FractionOfOne Value, XLColor Color)> ParseGradientStop(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(FractionOfOne Value, XLColor Color)>();
        }

        var position = _reader.GetDouble("position");

        var color = ParseColor("color", _ns).Value;
        _reader.Close(elementName, ns);

        return Xpr.From(OnGradientStopParsed(color, position));
    }

    private Xpr ParseBorders(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        var border = new List<XLDifferentialBorderValue>();
        while (ParseBorder("border", _ns) is { IsSuccess: true} borderItem)
        {
            border.Add(borderItem.Value);
        }
        _reader.Close(elementName, ns);

        OnBordersParsed(border, count);
        return Xpr.Success();
    }

    partial void OnBordersParsed(List<XLDifferentialBorderValue> border, uint? count);

    private Xpr<XLDifferentialBorderValue> ParseBorder(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLDifferentialBorderValue>();
        }

        var diagonalUp = _reader.GetOptionalBool("diagonalUp");
        var diagonalDown = _reader.GetOptionalBool("diagonalDown");
        var outline = _reader.GetOptionalBool("outline") ?? true;

        var leftResult = ParseBorderPr("left", _ns);
        var left = leftResult.IsSuccess ? leftResult.Value : default(XLBorderLine?);
        var rightResult = ParseBorderPr("right", _ns);
        var right = rightResult.IsSuccess ? rightResult.Value : default(XLBorderLine?);
        var topResult = ParseBorderPr("top", _ns);
        var top = topResult.IsSuccess ? topResult.Value : default(XLBorderLine?);
        var bottomResult = ParseBorderPr("bottom", _ns);
        var bottom = bottomResult.IsSuccess ? bottomResult.Value : default(XLBorderLine?);
        var diagonalResult = ParseBorderPr("diagonal", _ns);
        var diagonal = diagonalResult.IsSuccess ? diagonalResult.Value : default(XLBorderLine?);
        var verticalResult = ParseBorderPr("vertical", _ns);
        var vertical = verticalResult.IsSuccess ? verticalResult.Value : default(XLBorderLine?);
        var horizontalResult = ParseBorderPr("horizontal", _ns);
        var horizontal = horizontalResult.IsSuccess ? horizontalResult.Value : default(XLBorderLine?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnBorderParsed(left, right, top, bottom, diagonal, vertical, horizontal, diagonalUp, diagonalDown, outline));
    }

    private Xpr<XLBorderLine> ParseBorderPr(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLBorderLine>();
        }

        var style = _reader.GetOptionalEnum<XLBorderStyleValues>("style") ?? XLBorderStyleValues.None;

        var colorResult = ParseColor("color", _ns);
        var color = colorResult.IsSuccess ? colorResult.Value : default(XLColor?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnBorderPrParsed(color, style));
    }

    private Xpr ParseCellStyleXfs(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        var xf = new List<(XLCellFormatValue Format, int? CellStyleXfId)>();
        while (ParseXf("xf", _ns) is { IsSuccess: true} xfItem)
        {
            xf.Add(xfItem.Value);
        }

        if (xf.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnCellStyleXfsParsed(xf, count);
        return Xpr.Success();
    }

    partial void OnCellStyleXfsParsed(List<(XLCellFormatValue Format, int? CellStyleXfId)> xf, uint? count);

    private Xpr<(XLCellFormatValue Format, int? CellStyleXfId)> ParseXf(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(XLCellFormatValue Format, int? CellStyleXfId)>();
        }

        var numFmtId = _reader.GetOptionalUInt("numFmtId");
        var fontId = _reader.GetOptionalUInt("fontId");
        var fillId = _reader.GetOptionalUInt("fillId");
        var borderId = _reader.GetOptionalUInt("borderId");
        var xfId = _reader.GetOptionalUInt("xfId");
        var quotePrefix = _reader.GetOptionalBool("quotePrefix") ?? false;
        var pivotButton = _reader.GetOptionalBool("pivotButton") ?? false;
        var applyNumberFormat = _reader.GetOptionalBool("applyNumberFormat");
        var applyFont = _reader.GetOptionalBool("applyFont");
        var applyFill = _reader.GetOptionalBool("applyFill");
        var applyBorder = _reader.GetOptionalBool("applyBorder");
        var applyAlignment = _reader.GetOptionalBool("applyAlignment");
        var applyProtection = _reader.GetOptionalBool("applyProtection");

        var alignmentResult = ParseCellAlignment("alignment", _ns);
        var alignment = alignmentResult.IsSuccess ? alignmentResult.Value : default(XLDifferentialAlignmentValue?);
        var protectionResult = ParseCellProtection("protection", _ns);
        var protection = protectionResult.IsSuccess ? protectionResult.Value : default(XLDifferentialProtectionValue?);
        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnXfParsed(alignment, protection, numFmtId, fontId, fillId, borderId, xfId, quotePrefix, pivotButton, applyNumberFormat, applyFont, applyFill, applyBorder, applyAlignment, applyProtection));
    }

    private Xpr<XLDifferentialAlignmentValue> ParseCellAlignment(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLDifferentialAlignmentValue>();
        }

        var horizontal = _reader.GetOptionalEnum<XLAlignmentHorizontalValues>("horizontal");
        var vertical = _reader.GetOptionalEnum<XLAlignmentVerticalValues>("vertical") ?? XLAlignmentVerticalValues.Bottom;
        var textRotation = _reader.GetOptionalUInt("textRotation");
        var wrapText = _reader.GetOptionalBool("wrapText");
        var indent = _reader.GetOptionalUInt("indent");
        var relativeIndent = _reader.GetOptionalInt("relativeIndent");
        var justifyLastLine = _reader.GetOptionalBool("justifyLastLine");
        var shrinkToFit = _reader.GetOptionalBool("shrinkToFit");
        var readingOrder = _reader.GetOptionalUInt("readingOrder");

        _reader.Close(elementName, ns);

        return Xpr.From(OnCellAlignmentParsed(horizontal, vertical, textRotation, wrapText, indent, relativeIndent, justifyLastLine, shrinkToFit, readingOrder));
    }

    private Xpr<XLDifferentialProtectionValue> ParseCellProtection(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLDifferentialProtectionValue>();
        }

        var locked = _reader.GetOptionalBool("locked");
        var hidden = _reader.GetOptionalBool("hidden");

        _reader.Close(elementName, ns);

        return Xpr.From(OnCellProtectionParsed(locked, hidden));
    }

    private Xpr<List<(XLCellFormatValue Format, int? CellStyleXfId)>> ParseCellXfs(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<List<(XLCellFormatValue Format, int? CellStyleXfId)>>();
        }

        var count = _reader.GetOptionalUInt("count");

        var xf = new List<(XLCellFormatValue Format, int? CellStyleXfId)>();
        while (ParseXf("xf", _ns) is { IsSuccess: true} xfItem)
        {
            xf.Add(xfItem.Value);
        }

        if (xf.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnCellXfsParsed(xf, count));
    }

    private Xpr<Dictionary<int, XLCellStyleValue>> ParseCellStyles(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<Dictionary<int, XLCellStyleValue>>();
        }

        var count = _reader.GetOptionalUInt("count");

        var cellStyle = new List<(int CellStyleXfId, XLCellStyleValue Style)>();
        while (ParseCellStyle("cellStyle", _ns) is { IsSuccess: true} cellStyleItem)
        {
            cellStyle.Add(cellStyleItem.Value);
        }

        if (cellStyle.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnCellStylesParsed(cellStyle, count));
    }

    private Xpr<(int CellStyleXfId, XLCellStyleValue Style)> ParseCellStyle(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(int CellStyleXfId, XLCellStyleValue Style)>();
        }

        var name = _reader.GetOptionalXString("name");
        var xfId = _reader.GetUInt("xfId");
        var builtinId = _reader.GetOptionalUInt("builtinId");
        var iLevel = _reader.GetOptionalUInt("iLevel");
        var hidden = _reader.GetOptionalBool("hidden");
        var customBuiltin = _reader.GetOptionalBool("customBuiltin");

        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnCellStyleParsed(name, xfId, builtinId, iLevel, hidden, customBuiltin));
    }

    private Xpr ParseDxfs(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        while (ParseDxf("dxf", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'dxf' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnDxfsParsed(count);
        return Xpr.Success();
    }

    partial void OnDxfsParsed(uint? count);

    private Xpr ParseDxf(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var fontResult = ParseFont("font", _ns);
        var font = fontResult.IsSuccess ? fontResult.Value : default(XLDifferentialFontValue?);
        var numFmtResult = ParseNumFmt("numFmt", _ns);
        var numFmt = numFmtResult.IsSuccess ? numFmtResult.Value : default((int NumFmtId, XLNumberFormat Format)?);
        var fillResult = ParseFill("fill", _ns);
        var fill = fillResult.IsSuccess ? fillResult.Value : default(XLFillFormatValue?);
        var alignmentResult = ParseCellAlignment("alignment", _ns);
        var alignment = alignmentResult.IsSuccess ? alignmentResult.Value : default(XLDifferentialAlignmentValue?);
        var borderResult = ParseBorder("border", _ns);
        var border = borderResult.IsSuccess ? borderResult.Value : default(XLDifferentialBorderValue?);
        var protectionResult = ParseCellProtection("protection", _ns);
        var protection = protectionResult.IsSuccess ? protectionResult.Value : default(XLDifferentialProtectionValue?);
        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        OnDxfParsed(font, numFmt, fill, alignment, border, protection);
        return Xpr.Success();
    }

    partial void OnDxfParsed(XLDifferentialFontValue? font, (int NumFmtId, XLNumberFormat Format)? numFmt, XLFillFormatValue? fill, XLDifferentialAlignmentValue? alignment, XLDifferentialBorderValue? border, XLDifferentialProtectionValue? protection);

    private Xpr ParseTableStyles(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");
        var defaultTableStyle = _reader.GetOptionalString("defaultTableStyle");
        var defaultPivotStyle = _reader.GetOptionalString("defaultPivotStyle");

        while (ParseTableStyle("tableStyle", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tableStyle' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnTableStylesParsed(count, defaultTableStyle, defaultPivotStyle);
        return Xpr.Success();
    }

    partial void OnTableStylesParsed(uint? count, string? defaultTableStyle, string? defaultPivotStyle);

    private Xpr ParseTableStyle(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var name = _reader.GetString("name");
        var pivot = _reader.GetOptionalBool("pivot") ?? true;
        var table = _reader.GetOptionalBool("table") ?? true;
        var count = _reader.GetOptionalUInt("count");

        while (ParseTableStyleElement("tableStyleElement", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tableStyleElement' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnTableStyleParsed(name, pivot, table, count);
        return Xpr.Success();
    }

    partial void OnTableStyleParsed(string name, bool pivot, bool table, uint? count);

    private Xpr ParseTableStyleElement(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var type = _reader.GetStringMappedValue("type", TableStyleTypeMap);
        var size = _reader.GetOptionalUInt("size") ?? 1;
        var dxfId = _reader.GetOptionalUInt("dxfId");

        _reader.Close(elementName, ns);

        OnTableStyleElementParsed(type, size, dxfId);
        return Xpr.Success();
    }

    partial void OnTableStyleElementParsed((XLTableStyleRegionValues?, XLPivotStyleRegionValues?) type, uint size, uint? dxfId);

    private Xpr ParseColors(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        if (ParseIndexedColors("indexedColors", _ns) is { IsSuccess: true })
        {
            // Optional element 'indexedColors' was present
        }
        if (ParseMRUColors("mruColors", _ns) is { IsSuccess: true })
        {
            // Optional element 'mruColors' was present
        }
        _reader.Close(elementName, ns);

        OnColorsParsed();
        return Xpr.Success();
    }

    partial void OnColorsParsed();

    private Xpr ParseIndexedColors(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var rgbColor = new List<uint>();
        while (ParseRgbColor("rgbColor", _ns) is { IsSuccess: true} rgbColorItem)
        {
            rgbColor.Add(rgbColorItem.Value);
        }

        if (rgbColor.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnIndexedColorsParsed(rgbColor);
        return Xpr.Success();
    }

    partial void OnIndexedColorsParsed(List<uint> rgbColor);

    private Xpr ParseMRUColors(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var color = new List<XLColor>();
        while (ParseColor("color", _ns) is { IsSuccess: true} colorItem)
        {
            color.Add(colorItem.Value);
        }

        if (color.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnMRUColorsParsed(color);
        return Xpr.Success();
    }

    partial void OnMRUColorsParsed(List<XLColor> color);

    private Xpr<uint> ParseRgbColor(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<uint>();
        }

        var rgb = _reader.GetOptionalUIntHex("rgb");

        _reader.Close(elementName, ns);

        return Xpr.From(OnRgbColorParsed(rgb));
    }
}
