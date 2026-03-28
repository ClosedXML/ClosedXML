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
        unchecked
        {
            var hash = Bold ? 1 : 0;
            hash = (hash * 397) ^ (Italic ? 1 : 0);
            hash = (hash * 397) ^ (int)Underline;
            hash = (hash * 397) ^ (Strikethrough ? 1 : 0);
            hash = (hash * 397) ^ (int)VerticalAlignment;
            hash = (hash * 397) ^ (Shadow ? 1 : 0);

            // FontSize as long bits (BitConverter ensures stable bit pattern)
            hash = (hash * 397) ^ BitConverter.DoubleToInt64Bits(FontSize).GetHashCode();

            hash = (hash * 397) ^ FontColor.GetHashCode();

            // FontName with OrdinalIgnoreCase hash
            if (!string.IsNullOrEmpty(FontName))
                hash = (hash * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(FontName);

            hash = (hash * 397) ^ (int)FontFamilyNumbering;
            hash = (hash * 397) ^ (int)FontCharSet;
            hash = (hash * 397) ^ (int)FontScheme;

            return hash;
        }
    }

    public override string ToString()
    {
        return $"{FontName} {FontSize}pt {FontColor} " +
               (Bold ? "Bold" : "") + (Italic ? "Italic" : "") + (Strikethrough ? "Strikethrough" : "") +
               (Underline == XLFontUnderlineValues.None ? "" : Underline.ToString()) +
               $"{FontFamilyNumbering} {FontCharSet} {FontScheme}";
    }
}
