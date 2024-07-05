using System;

namespace ClosedXML.Excel;

internal readonly record struct XLFontKey
{
    public required bool Bold { get; init; }

    public required bool Italic { get; init; }

    public required XLFontUnderlineValues Underline { get; init; }

    public required bool Strikethrough { get; init; }

    public required XLFontVerticalTextAlignmentValues VerticalAlignment { get; init; }

    public required bool Shadow { get; init; }

    public required double FontSize { get; init; }

    public required XLColorKey FontColor { get; init; }

    public required string FontName { get; init; }

    public required XLFontFamilyNumberingValues FontFamilyNumbering { get; init; }

    public required XLFontCharSet FontCharSet { get; init; }

    public required XLFontScheme FontScheme { get; init; }

    public bool Equals(XLFontKey other)
    {
        return
            Bold == other.Bold
            && Italic == other.Italic
            && Underline == other.Underline
            && Strikethrough == other.Strikethrough
            && VerticalAlignment == other.VerticalAlignment
            && Shadow == other.Shadow
            && FontSize.Equals(other.FontSize)
            && FontColor == other.FontColor
            && FontFamilyNumbering == other.FontFamilyNumbering
            && FontCharSet == other.FontCharSet
            && FontScheme == other.FontScheme
            && string.Equals(FontName, other.FontName, StringComparison.OrdinalIgnoreCase);
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Bold);
        hash.Add(Italic);
        hash.Add(Underline);
        hash.Add(Strikethrough);
        hash.Add(VerticalAlignment);
        hash.Add(Shadow);
        hash.Add(FontSize);
        hash.Add(FontColor);
        hash.Add(FontName, StringComparer.OrdinalIgnoreCase);
        hash.Add(FontFamilyNumbering);
        hash.Add(FontCharSet);
        hash.Add(FontScheme);
        return hash.ToHashCode();
    }

    public override string ToString()
    {
        return $"{FontName} {FontSize}pt {FontColor} " +
               (Bold ? "Bold" : "") + (Italic ? "Italic" : "") + (Strikethrough ? "Strikethrough" : "") +
               (Underline == XLFontUnderlineValues.None ? "" : Underline.ToString()) +
               $"{FontFamilyNumbering} {FontCharSet} {FontScheme}";
    }
}
