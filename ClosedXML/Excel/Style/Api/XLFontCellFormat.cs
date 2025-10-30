using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// </summary>
internal sealed partial class XLFontCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLFontCellFormat(XLCellFormat parent)
    {
        _parent = parent;
    }

    internal XLFontName Name
    {
        get => Resolve(static x => x.Font.Name);
        set => Modify(static (font, fontName) => font with { Name = fontName }, value);
    }

    internal XLFontCharSet Charset
    {
        get => Resolve(static x => x.Font.Charset);
        set => Modify(static (font, charset) => font with { Charset = charset }, value);
    }

    internal XLFontFamilyNumberingValues Family
    {
        get => Resolve(static x => x.Font.Family);
        set => Modify(static (font, family) => font with { Family = family }, value);
    }

    internal bool Bold
    {
        get => Resolve(static x => x.Font.Bold);
        set => Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    internal bool Italic
    {
        get => Resolve(static x => x.Font.Italic);
        set => Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    internal bool Strikethrough
    {
        get => Resolve(static x => x.Font.Strikethrough);
        set => Modify(static (font, strikethrough) => font with { Strikethrough = strikethrough }, value);
    }

    internal bool Outline
    {
        get => Resolve(static x => x.Font.Outline);
        set => Modify(static (font, outline) => font with { Outline = outline }, value);
    }

    internal bool Shadow
    {
        get => Resolve(static x => x.Font.Shadow);
        set => Modify(static (font, shadow) => font with { Shadow = shadow }, value);
    }

    internal XLColor Color
    {
        get => Resolve(static x => x.Font.Color);
        set => Modify(static (font, color) => font with { Color = color }, value);
    }

    internal XLFontSize Size
    {
        get => Resolve(static x => x.Font.Size);
        set => Modify(static (font, size) => font with { Size = size }, value);
    }

    internal XLFontUnderlineValues Underline
    {
        get => Resolve(static x => x.Font.Underline);
        set => Modify(static (font, underline) => font with { Underline = underline }, value);
    }

    internal XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => Resolve(static x => x.Font.VerticalAlignment);
        set => Modify(static (font, verticalAlignment) => font with { VerticalAlignment = verticalAlignment }, value);
    }

    internal XLFontScheme Scheme
    {
        get => Resolve(static x => x.Font.Scheme);
        set => Modify(static (font, scheme) => font with { Scheme = scheme }, value);
    }

    private T Resolve<T>(Func<XLCellFormatValue, T> selector)
    {
        return _parent.Resolve(selector);
    }

    private void Modify<TProperty>(Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont, TProperty value)
    {
        _parent.ModifyFont(modifyFont, value);
    }

    /// <summary>
    /// A helper method to set all font properties at once (e.g, <c>someStyle.Font = otherStyle.Font</c>).
    /// </summary>
    internal void SetFont(IXLFont value)
    {
        _parent.ModifyFont(static (font, value) => font with
        {
            Bold = value.Bold,
            Italic = value.Italic,
            Underline = value.Underline,
            Strikethrough = value.Strikethrough,
            VerticalAlignment = value.VerticalAlignment,
            Shadow = value.Shadow,
            Size = XLFontSize.FromPoints(value.FontSize),
            Color = value.FontColor,
            Name = value.FontName,
            Family = value.FontFamilyNumbering,
            Charset = value.FontCharSet,
            Scheme = value.FontScheme
        }, value);
    }
}
