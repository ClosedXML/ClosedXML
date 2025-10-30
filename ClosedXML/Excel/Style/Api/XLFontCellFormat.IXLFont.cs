using System;

namespace ClosedXML.Excel;

internal sealed partial class XLFontCellFormat : IXLFont
{
    bool IXLFontBase.Bold
    {
        get => Bold;
        set => Bold = value;
    }

    bool IXLFontBase.Italic
    {
        get => Italic;
        set => Italic = value;
    }

    XLFontUnderlineValues IXLFontBase.Underline
    {
        get => Underline;
        set => Underline = value;
    }

    bool IXLFontBase.Strikethrough
    {
        get => Strikethrough;
        set => Strikethrough = value;
    }

    XLFontVerticalTextAlignmentValues IXLFontBase.VerticalAlignment
    {
        get => VerticalAlignment;
        set => VerticalAlignment = value;
    }

    bool IXLFontBase.Shadow
    {
        get => Shadow;
        set => Shadow = value;
    }

    double IXLFontBase.FontSize
    {
        get => Size.Points;
        set => Size = XLFontSize.FromPoints(value);
    }

    XLColor IXLFontBase.FontColor
    {
        get => Color;
        set => Color = value;
    }

    string IXLFontBase.FontName
    {
        get => Name.Text;
        set => Name = value;
    }

    XLFontFamilyNumberingValues IXLFontBase.FontFamilyNumbering
    {
        get => Family;
        set => Family = value;
    }

    XLFontCharSet IXLFontBase.FontCharSet
    {
        get => Charset;
        set => Charset = value;
    }

    XLFontScheme IXLFontBase.FontScheme
    {
        get => Scheme;
        set => Scheme = value;
    }

    bool IEquatable<IXLFont>.Equals(IXLFont? other)
    {
        if (other is null)
            return false;

        if (Bold != other.Bold)
            return false;

        if (Italic != other.Italic)
            return false;

        if (Underline != other.Underline)
            return false;

        if (Strikethrough != other.Strikethrough)
            return false;

        if (VerticalAlignment != other.VerticalAlignment)
            return false;

        if (Shadow != other.Shadow)
            return false;

        if (!Size.Points.Equals(other.FontSize))
            return false;

        if (Color != other.FontColor)
            return false;

        if (Name != other.FontName)
            return false;

        if (Family != other.FontFamilyNumbering)
            return false;

        if (Charset != other.FontCharSet)
            return false;

        if (Scheme != other.FontScheme)
            return false;

        return true;
    }

    IXLStyle IXLFont.SetBold()
    {
        return (this as IXLFont).SetBold(true);
    }

    IXLStyle IXLFont.SetBold(bool value)
    {
        Bold = value;
        return _parent;
    }

    IXLStyle IXLFont.SetItalic()
    {
        return (this as IXLFont).SetItalic(true);
    }

    IXLStyle IXLFont.SetItalic(bool value)
    {
        Italic = value;
        return _parent;
    }

    IXLStyle IXLFont.SetUnderline()
    {
        return (this as IXLFont).SetUnderline(XLFontUnderlineValues.Single);
    }

    IXLStyle IXLFont.SetUnderline(XLFontUnderlineValues value)
    {
        Underline = value;
        return _parent;
    }

    IXLStyle IXLFont.SetStrikethrough()
    {
        return (this as IXLFont).SetStrikethrough(true);
    }

    IXLStyle IXLFont.SetStrikethrough(bool value)
    {
        Strikethrough = value;
        return _parent;
    }

    IXLStyle IXLFont.SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        VerticalAlignment = value;
        return _parent;
    }

    IXLStyle IXLFont.SetShadow()
    {
        return (this as IXLFont).SetShadow(true);
    }

    IXLStyle IXLFont.SetShadow(bool value)
    {
        Shadow = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontSize(double value)
    {
        (this as IXLFont).FontSize = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontColor(XLColor value)
    {
        Color = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontName(string value)
    {
        Name = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        Family = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontCharSet(XLFontCharSet value)
    {
        Charset = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontScheme(XLFontScheme value)
    {
        Scheme = value;
        return _parent;
    }
}
