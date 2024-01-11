using System;

namespace ClosedXML.Excel
{
    internal struct XLFontKey : IEquatable<XLFontKey>
    {
        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public XLFontUnderlineValues Underline { get; set; }

        public bool Strikethrough { get; set; }

        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

        public bool Shadow { get; set; }

        public double FontSize { get; set; }

        public XLColorKey FontColor { get; set; }

        public string FontName { get; set; }

        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

        public XLFontCharSet FontCharSet { get; set; }

        public XLFontScheme FontScheme { get; set; }

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

        public override bool Equals(object obj)
        {
            if (obj is XLFontKey)
                return Equals((XLFontKey)obj);
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            var hashCode = 1158783753;
            hashCode = hashCode * -1521134295 + Bold.GetHashCode();
            hashCode = hashCode * -1521134295 + Italic.GetHashCode();
            hashCode = hashCode * -1521134295 + (int)Underline;
            hashCode = hashCode * -1521134295 + Strikethrough.GetHashCode();
            hashCode = hashCode * -1521134295 + (int)VerticalAlignment;
            hashCode = hashCode * -1521134295 + Shadow.GetHashCode();
            hashCode = hashCode * -1521134295 + FontSize.GetHashCode();
            hashCode = hashCode * -1521134295 + FontColor.GetHashCode();
            hashCode = hashCode * -1521134295 + StringComparer.OrdinalIgnoreCase.GetHashCode(FontName);
            hashCode = hashCode * -1521134295 + (int)FontFamilyNumbering;
            hashCode = hashCode * -1521134295 + (int)FontCharSet;
            hashCode = hashCode * -1521134295 + (int)FontScheme;
            return hashCode;
        }

        public override string ToString()
        {
            return $"{FontName} {FontSize}pt {FontColor} " +
                   (Bold ? "Bold" : "") + (Italic ? "Italic" : "") + (Strikethrough ? "Strikethrough" : "") +
                   (Underline == XLFontUnderlineValues.None ? "" : Underline.ToString()) +
                   $"{FontFamilyNumbering} {FontCharSet} {FontScheme}";
        }

        public static bool operator ==(XLFontKey left, XLFontKey right) => left.Equals(right);

        public static bool operator !=(XLFontKey left, XLFontKey right) => !(left.Equals(right));
    }
}
