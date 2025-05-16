#nullable enable

using System.Collections.Generic;
using ClosedXML.IO;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel.IO;

internal partial class StylesReader
{
    private void ParseNumFmts(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        var numFmt = new List<(int NumFmtId, string FormatCode)>();
        while (_reader.TryOpen("numFmt", _ns))
        {
            numFmt.Add(ParseNumFmt("numFmt"));
        }
        _reader.Close(elementName, _ns);
        OnNumFmtsParsed(numFmt, count);
    }

    partial void OnNumFmtsParsed(List<(int NumFmtId, string FormatCode)> numFmt, uint? count);

    private (int NumFmtId, string FormatCode) ParseNumFmt(string elementName)
    {
        var numFmtId = _reader.GetUInt("numFmtId");
        var formatCode = _reader.GetXString("formatCode");
        _reader.Close(elementName, _ns);
        return OnNumFmtParsed(numFmtId, formatCode);
    }

    private void ParseFonts(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        while (_reader.TryOpen("font", _ns))
        {
            ParseFont("font");
        }
        _reader.Close(elementName, _ns);
        OnFontsParsed(count);
    }

    partial void OnFontsParsed(uint? count);

    private void ParseFills(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        while (_reader.TryOpen("fill", _ns))
        {
            ParseFill("fill");
        }
        _reader.Close(elementName, _ns);
        OnFillsParsed(count);
    }

    partial void OnFillsParsed(uint? count);

    private void ParseFill(string elementName)
    {
        XLFillFormat? patternFill = null;
        XLFillFormat? gradientFill = null;
        if (_reader.TryOpen("patternFill", _ns))
        {
            patternFill = ParsePatternFill("patternFill");
        }
        else if (_reader.TryOpen("gradientFill", _ns))
        {
            gradientFill = ParseGradientFill("gradientFill");
        }
        _reader.Close(elementName, _ns);
        OnFillParsed(patternFill, gradientFill);
    }

    partial void OnFillParsed(XLFillFormat? patternFill, XLFillFormat? gradientFill);

    private XLFillFormat ParsePatternFill(string elementName)
    {
        var patternType = _reader.GetOptionalEnum<XLFillPatternValues>("patternType");
        XLColor? fgColor = default;
        if (_reader.TryOpen("fgColor", _ns))
        {
            fgColor = ParseColor("fgColor");
        }
        XLColor? bgColor = default;
        if (_reader.TryOpen("bgColor", _ns))
        {
            bgColor = ParseColor("bgColor");
        }
        _reader.Close(elementName, _ns);
        return OnPatternFillParsed(fgColor, bgColor, patternType);
    }

    private XLFillFormat ParseGradientFill(string elementName)
    {
        var type = _reader.GetOptionalEnum<XLGradientType>("type") ?? XLGradientType.Linear;
        var degree = _reader.GetOptionalDouble("degree") ?? 0;
        var left = _reader.GetOptionalDouble("left") ?? 0;
        var right = _reader.GetOptionalDouble("right") ?? 0;
        var top = _reader.GetOptionalDouble("top") ?? 0;
        var bottom = _reader.GetOptionalDouble("bottom") ?? 0;
        var stop = new List<(FractionOfOne Value, XLColor Color)>();
        while (_reader.TryOpen("stop", _ns))
        {
            stop.Add(ParseGradientStop("stop"));
        }
        _reader.Close(elementName, _ns);
        return OnGradientFillParsed(stop, type, degree, left, right, top, bottom);
    }

    private (FractionOfOne Value, XLColor Color) ParseGradientStop(string elementName)
    {
        var position = _reader.GetDouble("position");
        _reader.Open("color", _ns);
        var color = ParseColor("color");
        _reader.Close(elementName, _ns);
        return OnGradientStopParsed(color, position);
    }

    private void ParseBorders(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        while (_reader.TryOpen("border", _ns))
        {
            ParseBorder("border");
        }
        _reader.Close(elementName, _ns);
        OnBordersParsed(count);
    }

    partial void OnBordersParsed(uint? count);

    private void ParseBorder(string elementName)
    {
        var diagonalUp = _reader.GetOptionalBool("diagonalUp");
        var diagonalDown = _reader.GetOptionalBool("diagonalDown");
        var outline = _reader.GetOptionalBool("outline") ?? true;
        XLBorderLine? left = default;
        if (_reader.TryOpen("left", _ns))
        {
            left = ParseBorderPr("left");
        }
        XLBorderLine? right = default;
        if (_reader.TryOpen("right", _ns))
        {
            right = ParseBorderPr("right");
        }
        XLBorderLine? top = default;
        if (_reader.TryOpen("top", _ns))
        {
            top = ParseBorderPr("top");
        }
        XLBorderLine? bottom = default;
        if (_reader.TryOpen("bottom", _ns))
        {
            bottom = ParseBorderPr("bottom");
        }
        XLBorderLine? diagonal = default;
        if (_reader.TryOpen("diagonal", _ns))
        {
            diagonal = ParseBorderPr("diagonal");
        }
        XLBorderLine? vertical = default;
        if (_reader.TryOpen("vertical", _ns))
        {
            vertical = ParseBorderPr("vertical");
        }
        XLBorderLine? horizontal = default;
        if (_reader.TryOpen("horizontal", _ns))
        {
            horizontal = ParseBorderPr("horizontal");
        }
        _reader.Close(elementName, _ns);
        OnBorderParsed(left, right, top, bottom, diagonal, vertical, horizontal, diagonalUp, diagonalDown, outline);
    }

    partial void OnBorderParsed(XLBorderLine? left, XLBorderLine? right, XLBorderLine? top, XLBorderLine? bottom, XLBorderLine? diagonal, XLBorderLine? vertical, XLBorderLine? horizontal, bool? diagonalUp, bool? diagonalDown, bool outline);

    private XLBorderLine ParseBorderPr(string elementName)
    {
        var style = _reader.GetOptionalEnum<XLBorderStyleValues>("style") ?? XLBorderStyleValues.None;
        XLColor? color = default;
        if (_reader.TryOpen("color", _ns))
        {
            color = ParseColor("color");
        }
        _reader.Close(elementName, _ns);
        return OnBorderPrParsed(color, style);
    }

    private void ParseCellStyleXfs(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        var xf = new List<(XLCellFormat Format, int? CellStyleXfId)>();
        _reader.Open("xf", _ns);
        do
        {
            xf.Add(ParseXf("xf"));
        }
        while (_reader.TryOpen("xf", _ns));
        _reader.Close(elementName, _ns);
        OnCellStyleXfsParsed(xf, count);
    }

    partial void OnCellStyleXfsParsed(List<(XLCellFormat Format, int? CellStyleXfId)> xf, uint? count);

    private (XLCellFormat Format, int? CellStyleXfId) ParseXf(string elementName)
    {
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
        XLAlignmentFormat? alignment = default;
        if (_reader.TryOpen("alignment", _ns))
        {
            alignment = ParseCellAlignment("alignment");
        }
        XLProtectionFormat? protection = default;
        if (_reader.TryOpen("protection", _ns))
        {
            protection = ParseCellProtection("protection");
        }
        if (_reader.TryOpen("extLst", _ns))
        {
            ParseExtensionList("extLst");
        }
        _reader.Close(elementName, _ns);
        return OnXfParsed(alignment, protection, numFmtId, fontId, fillId, borderId, xfId, quotePrefix, pivotButton, applyNumberFormat, applyFont, applyFill, applyBorder, applyAlignment, applyProtection);
    }

    private XLAlignmentFormat ParseCellAlignment(string elementName)
    {
        var horizontal = _reader.GetOptionalEnum<XLAlignmentHorizontalValues>("horizontal");
        var vertical = _reader.GetOptionalEnum<XLAlignmentVerticalValues>("vertical") ?? XLAlignmentVerticalValues.Bottom;
        var textRotation = _reader.GetOptionalUInt("textRotation");
        var wrapText = _reader.GetOptionalBool("wrapText");
        var indent = _reader.GetOptionalUInt("indent");
        var relativeIndent = _reader.GetOptionalInt("relativeIndent");
        var justifyLastLine = _reader.GetOptionalBool("justifyLastLine");
        var shrinkToFit = _reader.GetOptionalBool("shrinkToFit");
        var readingOrder = _reader.GetOptionalUInt("readingOrder");
        _reader.Close(elementName, _ns);
        return OnCellAlignmentParsed(horizontal, vertical, textRotation, wrapText, indent, relativeIndent, justifyLastLine, shrinkToFit, readingOrder);
    }

    private XLProtectionFormat ParseCellProtection(string elementName)
    {
        var locked = _reader.GetOptionalBool("locked");
        var hidden = _reader.GetOptionalBool("hidden");
        _reader.Close(elementName, _ns);
        return OnCellProtectionParsed(locked, hidden);
    }

    private List<(XLCellFormat Format, int? CellStyleXfId)> ParseCellXfs(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        var xf = new List<(XLCellFormat Format, int? CellStyleXfId)>();
        _reader.Open("xf", _ns);
        do
        {
            xf.Add(ParseXf("xf"));
        }
        while (_reader.TryOpen("xf", _ns));
        _reader.Close(elementName, _ns);
        return OnCellXfsParsed(xf, count);
    }

    private Dictionary<int, XLCellStyle> ParseCellStyles(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        var cellStyle = new List<(int CellStyleXfId, XLCellStyle Style)>();
        _reader.Open("cellStyle", _ns);
        do
        {
            cellStyle.Add(ParseCellStyle("cellStyle"));
        }
        while (_reader.TryOpen("cellStyle", _ns));
        _reader.Close(elementName, _ns);
        return OnCellStylesParsed(cellStyle, count);
    }

    private (int CellStyleXfId, XLCellStyle Style) ParseCellStyle(string elementName)
    {
        var name = _reader.GetOptionalXString("name");
        var xfId = _reader.GetUInt("xfId");
        var builtinId = _reader.GetOptionalUInt("builtinId");
        var iLevel = _reader.GetOptionalUInt("iLevel");
        var hidden = _reader.GetOptionalBool("hidden");
        var customBuiltin = _reader.GetOptionalBool("customBuiltin");
        if (_reader.TryOpen("extLst", _ns))
        {
            ParseExtensionList("extLst");
        }
        _reader.Close(elementName, _ns);
        return OnCellStyleParsed(name, xfId, builtinId, iLevel, hidden, customBuiltin);
    }

    private void ParseDxfs(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        while (_reader.TryOpen("dxf", _ns))
        {
            ParseDxf("dxf");
        }
        _reader.Close(elementName, _ns);
        OnDxfsParsed(count);
    }

    partial void OnDxfsParsed(uint? count);

    private void ParseDxf(string elementName)
    {
        if (_reader.TryOpen("font", _ns))
        {
            ParseFont("font");
        }
        (int NumFmtId, string FormatCode)? numFmt = default;
        if (_reader.TryOpen("numFmt", _ns))
        {
            numFmt = ParseNumFmt("numFmt");
        }
        if (_reader.TryOpen("fill", _ns))
        {
            ParseFill("fill");
        }
        XLAlignmentFormat? alignment = default;
        if (_reader.TryOpen("alignment", _ns))
        {
            alignment = ParseCellAlignment("alignment");
        }
        if (_reader.TryOpen("border", _ns))
        {
            ParseBorder("border");
        }
        XLProtectionFormat? protection = default;
        if (_reader.TryOpen("protection", _ns))
        {
            protection = ParseCellProtection("protection");
        }
        if (_reader.TryOpen("extLst", _ns))
        {
            ParseExtensionList("extLst");
        }
        _reader.Close(elementName, _ns);
        OnDxfParsed(numFmt, alignment, protection);
    }

    partial void OnDxfParsed((int NumFmtId, string FormatCode)? numFmt, XLAlignmentFormat? alignment, XLProtectionFormat? protection);

    private void ParseTableStyles(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        var defaultTableStyle = _reader.GetOptionalString("defaultTableStyle");
        var defaultPivotStyle = _reader.GetOptionalString("defaultPivotStyle");
        while (_reader.TryOpen("tableStyle", _ns))
        {
            ParseTableStyle("tableStyle");
        }
        _reader.Close(elementName, _ns);
        OnTableStylesParsed(count, defaultTableStyle, defaultPivotStyle);
    }

    partial void OnTableStylesParsed(uint? count, string? defaultTableStyle, string? defaultPivotStyle);

    private void ParseTableStyle(string elementName)
    {
        var name = _reader.GetString("name");
        var pivot = _reader.GetOptionalBool("pivot") ?? true;
        var table = _reader.GetOptionalBool("table") ?? true;
        var count = _reader.GetOptionalUInt("count");
        while (_reader.TryOpen("tableStyleElement", _ns))
        {
            ParseTableStyleElement("tableStyleElement");
        }
        _reader.Close(elementName, _ns);
        OnTableStyleParsed(name, pivot, table, count);
    }

    partial void OnTableStyleParsed(string name, bool pivot, bool table, uint? count);

    private void ParseTableStyleElement(string elementName)
    {
        var type = _reader.GetEnum<XLTableStyleType>("type");
        var size = _reader.GetOptionalUInt("size") ?? 1;
        var dxfId = _reader.GetOptionalUInt("dxfId");
        _reader.Close(elementName, _ns);
        OnTableStyleElementParsed(type, size, dxfId);
    }

    partial void OnTableStyleElementParsed(XLTableStyleType type, uint size, uint? dxfId);

    private void ParseColors(string elementName)
    {
        if (_reader.TryOpen("indexedColors", _ns))
        {
            ParseIndexedColors("indexedColors");
        }
        if (_reader.TryOpen("mruColors", _ns))
        {
            ParseMRUColors("mruColors");
        }
        _reader.Close(elementName, _ns);
        OnColorsParsed();
    }

    partial void OnColorsParsed();

    private void ParseIndexedColors(string elementName)
    {
        _reader.Open("rgbColor", _ns);
        do
        {
            ParseRgbColor("rgbColor");
        }
        while (_reader.TryOpen("rgbColor", _ns));
        _reader.Close(elementName, _ns);
        OnIndexedColorsParsed();
    }

    partial void OnIndexedColorsParsed();

    private void ParseMRUColors(string elementName)
    {
        var color = new List<XLColor>();
        _reader.Open("color", _ns);
        do
        {
            color.Add(ParseColor("color"));
        }
        while (_reader.TryOpen("color", _ns));
        _reader.Close(elementName, _ns);
        OnMRUColorsParsed(color);
    }

    partial void OnMRUColorsParsed(List<XLColor> color);

    private void ParseRgbColor(string elementName)
    {
        var rgb = _reader.GetOptionalUIntHex("rgb");
        _reader.Close(elementName, _ns);
        OnRgbColorParsed(rgb);
    }

    partial void OnRgbColorParsed(uint? rgb);
}
