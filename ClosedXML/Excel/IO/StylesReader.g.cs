#nullable enable

using System.Collections.Generic;
using ClosedXML.IO;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel.IO;

internal partial class StylesReader
{
    private void ParseStylesheet(string elementName)
    {
        if (_reader.TryOpen("numFmts", _ns))
        {
            ParseNumFmts("numFmts");
        }
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
        if (_reader.TryOpen("cellXfs", _ns))
        {
            ParseCellXfs("cellXfs");
        }
        if (_reader.TryOpen("cellStyles", _ns))
        {
            ParseCellStyles("cellStyles");
        }
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
        OnStylesheetParsed();
    }

    partial void OnStylesheetParsed();

    private void ParseNumFmts(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        while (_reader.TryOpen("numFmt", _ns))
        {
            ParseNumFmt("numFmt");
        }
        _reader.Close(elementName, _ns);
        OnNumFmtsParsed(count);
    }

    partial void OnNumFmtsParsed(uint? count);

    private void ParseNumFmt(string elementName)
    {
        var numFmtId = _reader.GetUInt("numFmtId");
        var formatCode = _reader.GetXString("formatCode");
        _reader.Close(elementName, _ns);
        OnNumFmtParsed(numFmtId, formatCode);
    }

    partial void OnNumFmtParsed(uint numFmtId, string formatCode);

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
        if (_reader.TryOpen("patternFill", _ns))
        {
            ParsePatternFill("patternFill");
        }
        else if (_reader.TryOpen("gradientFill", _ns))
        {
            ParseGradientFill("gradientFill");
        }
        _reader.Close(elementName, _ns);
        OnFillParsed();
    }

    partial void OnFillParsed();

    private void ParsePatternFill(string elementName)
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
        OnPatternFillParsed(fgColor, bgColor, patternType);
    }

    partial void OnPatternFillParsed(XLColor? fgColor, XLColor? bgColor, XLFillPatternValues? patternType);

    private void ParseGradientFill(string elementName)
    {
        var type = _reader.GetOptionalEnum<XLGradientType>("type") ?? XLGradientType.Linear;
        var degree = _reader.GetOptionalDouble("degree") ?? 0;
        var left = _reader.GetOptionalDouble("left") ?? 0;
        var right = _reader.GetOptionalDouble("right") ?? 0;
        var top = _reader.GetOptionalDouble("top") ?? 0;
        var bottom = _reader.GetOptionalDouble("bottom") ?? 0;
        while (_reader.TryOpen("stop", _ns))
        {
            ParseGradientStop("stop");
        }
        _reader.Close(elementName, _ns);
        OnGradientFillParsed(type, degree, left, right, top, bottom);
    }

    partial void OnGradientFillParsed(XLGradientType type, double degree, double left, double right, double top, double bottom);

    private void ParseGradientStop(string elementName)
    {
        var position = _reader.GetDouble("position");
        _reader.Open("color", _ns);
        var color = ParseColor("color");
        _reader.Close(elementName, _ns);
        OnGradientStopParsed(color, position);
    }

    partial void OnGradientStopParsed(XLColor color, double position);

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
        if (_reader.TryOpen("left", _ns))
        {
            ParseBorderPr("left");
        }
        if (_reader.TryOpen("right", _ns))
        {
            ParseBorderPr("right");
        }
        if (_reader.TryOpen("top", _ns))
        {
            ParseBorderPr("top");
        }
        if (_reader.TryOpen("bottom", _ns))
        {
            ParseBorderPr("bottom");
        }
        if (_reader.TryOpen("diagonal", _ns))
        {
            ParseBorderPr("diagonal");
        }
        if (_reader.TryOpen("vertical", _ns))
        {
            ParseBorderPr("vertical");
        }
        if (_reader.TryOpen("horizontal", _ns))
        {
            ParseBorderPr("horizontal");
        }
        _reader.Close(elementName, _ns);
        OnBorderParsed(diagonalUp, diagonalDown, outline);
    }

    partial void OnBorderParsed(bool? diagonalUp, bool? diagonalDown, bool outline);

    private void ParseBorderPr(string elementName)
    {
        var style = _reader.GetOptionalEnum<XLBorderStyleValues>("style") ?? XLBorderStyleValues.None;
        XLColor? color = default;
        if (_reader.TryOpen("color", _ns))
        {
            color = ParseColor("color");
        }
        _reader.Close(elementName, _ns);
        OnBorderPrParsed(color, style);
    }

    partial void OnBorderPrParsed(XLColor? color, XLBorderStyleValues style);

    private void ParseCellStyleXfs(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        _reader.Open("xf", _ns);
        do
        {
            ParseXf("xf");
        }
        while (_reader.TryOpen("xf", _ns));
        _reader.Close(elementName, _ns);
        OnCellStyleXfsParsed(count);
    }

    partial void OnCellStyleXfsParsed(uint? count);

    private void ParseXf(string elementName)
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
        if (_reader.TryOpen("alignment", _ns))
        {
            ParseCellAlignment("alignment");
        }
        if (_reader.TryOpen("protection", _ns))
        {
            ParseCellProtection("protection");
        }
        if (_reader.TryOpen("extLst", _ns))
        {
            ParseExtensionList("extLst");
        }
        _reader.Close(elementName, _ns);
        OnXfParsed(numFmtId, fontId, fillId, borderId, xfId, quotePrefix, pivotButton, applyNumberFormat, applyFont, applyFill, applyBorder, applyAlignment, applyProtection);
    }

    partial void OnXfParsed(uint? numFmtId, uint? fontId, uint? fillId, uint? borderId, uint? xfId, bool quotePrefix, bool pivotButton, bool? applyNumberFormat, bool? applyFont, bool? applyFill, bool? applyBorder, bool? applyAlignment, bool? applyProtection);

    private void ParseCellAlignment(string elementName)
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
        OnCellAlignmentParsed(horizontal, vertical, textRotation, wrapText, indent, relativeIndent, justifyLastLine, shrinkToFit, readingOrder);
    }

    partial void OnCellAlignmentParsed(XLAlignmentHorizontalValues? horizontal, XLAlignmentVerticalValues vertical, uint? textRotation, bool? wrapText, uint? indent, int? relativeIndent, bool? justifyLastLine, bool? shrinkToFit, uint? readingOrder);

    private void ParseCellProtection(string elementName)
    {
        var locked = _reader.GetOptionalBool("locked");
        var hidden = _reader.GetOptionalBool("hidden");
        _reader.Close(elementName, _ns);
        OnCellProtectionParsed(locked, hidden);
    }

    partial void OnCellProtectionParsed(bool? locked, bool? hidden);

    private void ParseCellStyles(string elementName)
    {
        var count = _reader.GetOptionalUInt("count");
        _reader.Open("cellStyle", _ns);
        do
        {
            ParseCellStyle("cellStyle");
        }
        while (_reader.TryOpen("cellStyle", _ns));
        _reader.Close(elementName, _ns);
        OnCellStylesParsed(count);
    }

    partial void OnCellStylesParsed(uint? count);

    private void ParseCellStyle(string elementName)
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
        OnCellStyleParsed(name, xfId, builtinId, iLevel, hidden, customBuiltin);
    }

    partial void OnCellStyleParsed(string? name, uint xfId, uint? builtinId, uint? iLevel, bool? hidden, bool? customBuiltin);

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
        if (_reader.TryOpen("numFmt", _ns))
        {
            ParseNumFmt("numFmt");
        }
        if (_reader.TryOpen("fill", _ns))
        {
            ParseFill("fill");
        }
        if (_reader.TryOpen("alignment", _ns))
        {
            ParseCellAlignment("alignment");
        }
        if (_reader.TryOpen("border", _ns))
        {
            ParseBorder("border");
        }
        if (_reader.TryOpen("protection", _ns))
        {
            ParseCellProtection("protection");
        }
        if (_reader.TryOpen("extLst", _ns))
        {
            ParseExtensionList("extLst");
        }
        _reader.Close(elementName, _ns);
        OnDxfParsed();
    }

    partial void OnDxfParsed();

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
        _reader.Open("color", _ns);
        do
        {
            ParseColor("color");
        }
        while (_reader.TryOpen("color", _ns));
        _reader.Close(elementName, _ns);
        OnMRUColorsParsed();
    }

    partial void OnMRUColorsParsed();

    private void ParseRgbColor(string elementName)
    {
        var rgb = _reader.GetOptionalUIntHex("rgb");
        _reader.Close(elementName, _ns);
        OnRgbColorParsed(rgb);
    }

    partial void OnRgbColorParsed(uint? rgb);
}
