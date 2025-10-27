using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// </summary>
internal partial class XLFontCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLFontCellFormat(XLCellFormat parent)
    {
        _parent = parent;
    }

    public XLFontName Name
    {
        get => Resolve(static x => x.Font.Name);
        set => Modify(static (font, fontName) => font with { Name = fontName }, value);
    }

    public bool Bold
    {
        get => Resolve(static x => x.Font.Bold);
        set => Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    public bool Italic
    {
        get => Resolve(static x => x.Font.Italic);
        set => Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    public XLFontSize Size
    {
        get => Resolve(static x => x.Font.Size);
        set => Modify(static (font, size) => font with { Size = size }, value);
    }

    private T Resolve<T>(Func<XLCellFormatValue, T> selector)
    {
        return _parent.Resolve(selector);
    }

    private void Modify<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        _parent.ModifyFont(modifyFont, value);
    }
}
