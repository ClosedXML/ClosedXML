using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal partial class XLDxfFontFormat : IXLFont
{
    private readonly XLFontFormatValue _defaultFont = XLFontFormatValue.Default;

    bool IXLFontBase.Bold
    {
        get => Resolve(static font => font.Bold, _defaultFont.Bold);
        set => Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    bool IXLFontBase.Italic
    {
        get => Resolve(static font => font.Italic, _defaultFont.Italic);
        set => Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    XLFontUnderlineValues IXLFontBase.Underline
    {
        get => Resolve(static font => font.Underline, _defaultFont.Underline);
        set => Modify(static (font, underline) => font with { Underline = underline }, value);
    }

    bool IXLFontBase.Strikethrough
    {
        get => Resolve(static font => font.Strikethrough, _defaultFont.Strikethrough);
        set => Modify(static (font, strikethrough) => font with { Strikethrough = strikethrough }, value);
    }

    XLFontVerticalTextAlignmentValues IXLFontBase.VerticalAlignment
    {
        get => Resolve(static font => font.VerticalAlignment, _defaultFont.VerticalAlignment);
        set => Modify(static (font, vAlign) => font with { VerticalAlignment = vAlign }, value);
    }

    bool IXLFontBase.Shadow
    {
        get => Resolve(static font => font.Shadow, _defaultFont.Shadow);
        set => Modify(static (font, shadow) => font with { Shadow = shadow }, value);
    }

    double IXLFontBase.FontSize
    {
        get => Resolve(static font => font.Size?.Points, _defaultFont.Size.Points);
        set => Modify(static (font, size) => font with { Size = XLFontSize.FromPoints(size) }, value);
    }

    XLColor IXLFontBase.FontColor
    {
        get => Resolve(static font => font.Color, _defaultFont.Color);
        set => Modify(static (font, color) => font with { Color = color }, value);
    }

    string IXLFontBase.FontName
    {
        get => Resolve(static font => font.Name?.Text, _defaultFont.Name.Text);
        set => Modify(static (font, name) => font with { Name = name }, value);
    }

    XLFontFamilyNumberingValues IXLFontBase.FontFamilyNumbering
    {
        get => Resolve(static font => font.Family, _defaultFont.Family);
        set => Modify(static (font, family) => font with { Family = family }, value);
    }

    XLFontCharSet IXLFontBase.FontCharSet
    {
        get => Resolve(static font => font.Charset, _defaultFont.Charset);
        set => Modify(static (font, charset) => font with { Charset = charset }, value);
    }

    XLFontScheme IXLFontBase.FontScheme
    {
        get => Resolve(static font => font.Scheme, _defaultFont.Scheme);
        set => Modify(static (font, scheme) => font with { Scheme = scheme }, value);
    }

    IXLStyle IXLFont.SetBold()
    {
        return (this as IXLFont).SetBold(true);
    }

    IXLStyle IXLFont.SetBold(bool value)
    {
        (this as IXLFont).Bold = value;
        return _parent;
    }

    IXLStyle IXLFont.SetItalic()
    {
        return (this as IXLFont).SetItalic(true);
    }

    IXLStyle IXLFont.SetItalic(bool value)
    {
        (this as IXLFont).Italic = value;
        return _parent;
    }

    IXLStyle IXLFont.SetUnderline()
    {
        return (this as IXLFont).SetUnderline(XLFontUnderlineValues.Single);
    }

    IXLStyle IXLFont.SetUnderline(XLFontUnderlineValues value)
    {
        (this as IXLFont).Underline = value;
        return _parent;
    }

    IXLStyle IXLFont.SetStrikethrough()
    {
        return (this as IXLFont).SetStrikethrough(true);
    }

    IXLStyle IXLFont.SetStrikethrough(bool value)
    {
        (this as IXLFont).Strikethrough = value;
        return _parent;
    }

    IXLStyle IXLFont.SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        (this as IXLFont).VerticalAlignment = value;
        return _parent;
    }

    IXLStyle IXLFont.SetShadow()
    {
        return (this as IXLFont).SetShadow(true);
    }

    IXLStyle IXLFont.SetShadow(bool value)
    {
        (this as IXLFont).Shadow = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontSize(double value)
    {
        (this as IXLFont).FontSize = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontColor(XLColor value)
    {
        (this as IXLFont).FontColor = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontName(string value)
    {
        (this as IXLFont).FontName = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        (this as IXLFont).FontFamilyNumbering = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontCharSet(XLFontCharSet value)
    {
        (this as IXLFont).FontCharSet = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontScheme(XLFontScheme value)
    {
        (this as IXLFont).FontScheme = value;
        return _parent;
    }

    bool IEquatable<IXLFont>.Equals(IXLFont? other)
    {
        throw new NotSupportedException();
    }
}
