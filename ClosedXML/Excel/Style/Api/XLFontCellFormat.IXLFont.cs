using System;

namespace ClosedXML.Excel;

internal partial class XLFontCellFormat : IXLFont
{
    // TODO Styles: Implement remaining format properties by using IXLFont contract
    XLFontUnderlineValues IXLFontBase.Underline
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    bool IXLFontBase.Strikethrough
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    XLFontVerticalTextAlignmentValues IXLFontBase.VerticalAlignment
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    bool IXLFontBase.Shadow
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    double IXLFontBase.FontSize
    {
        get => Size.Points;
        set => Size = XLFontSize.FromPoints(value);
    }

    XLColor IXLFontBase.FontColor
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    string IXLFontBase.FontName
    {
        get => Name.Text;
        set => Name = value;
    }

    XLFontFamilyNumberingValues IXLFontBase.FontFamilyNumbering
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    XLFontCharSet IXLFontBase.FontCharSet
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    XLFontScheme IXLFontBase.FontScheme
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    bool IEquatable<IXLFont>.Equals(IXLFont? other)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetBold()
    {
        return ((IXLFont)this).SetBold(true);
    }

    IXLStyle IXLFont.SetBold(bool value)
    {
        Bold = value;
        return _parent;
    }

    IXLStyle IXLFont.SetItalic()
    {
        return ((IXLFont)this).SetItalic(true);
    }

    IXLStyle IXLFont.SetItalic(bool value)
    {
        Italic = value;
        return _parent;
    }

    IXLStyle IXLFont.SetUnderline()
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetUnderline(XLFontUnderlineValues value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetStrikethrough()
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetStrikethrough(bool value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetShadow()
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetShadow(bool value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetFontSize(double value)
    {
        ((IXLFont)this).FontSize = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontColor(XLColor value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetFontName(string value)
    {
        Name = value;
        return _parent;
    }

    IXLStyle IXLFont.SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetFontCharSet(XLFontCharSet value)
    {
        throw new NotImplementedException();
    }

    IXLStyle IXLFont.SetFontScheme(XLFontScheme value)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// A helper method to set all font properties at once (e.g, <c>someStyle.Font = otherStyle.Font</c>).
    /// </summary>
    internal void SetFont(IXLFont value)
    {
        throw new NotImplementedException();
    }
}
