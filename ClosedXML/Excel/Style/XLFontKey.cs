using System;

namespace ClosedXML.Excel
{
    using DocumentFormat.OpenXml.Office2010.PowerPoint;

    internal struct XLFontKey : IEquatable<XLFontKey>
    {
        private bool _bold;

        private bool _italic;

        private XLFontUnderlineValues _underline;

        private bool _strikethrough;

        private XLFontVerticalTextAlignmentValues _verticalAlignment;

        private bool _shadow;

        private double _fontSize;

        private XLColorKey _fontColor;

        private string _fontName;

        private XLFontFamilyNumberingValues _fontFamilyNumbering;

        private XLFontCharSet _fontCharSet;

        private int _cachedHashCode;

        public bool Bold
        {
            get { return _bold; }
            set
            {
                _bold = value;
                _cachedHashCode = 0;
            }
        }

        public bool Italic
        {
            get { return _italic; }
            set
            {
                _italic = value;
                _cachedHashCode = 0;
            }
        }

        public XLFontUnderlineValues Underline
        {
            get { return _underline; }
            set
            {
                _underline = value;
                _cachedHashCode = 0;
            }
        }

        public bool Strikethrough
        {
            get { return _strikethrough; }
            set
            {
                _strikethrough = value;
                _cachedHashCode = 0;
            }
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get { return _verticalAlignment; }
            set
            {
                _verticalAlignment = value;
                _cachedHashCode = 0;
            }
        }

        public bool Shadow
        {
            get { return _shadow; }
            set
            {
                _shadow = value;
                _cachedHashCode = 0;
            }
        }

        public double FontSize
        {
            get { return _fontSize; }
            set
            {
                _fontSize = value;
                _cachedHashCode = 0;
            }
        }

        public XLColorKey FontColor
        {
            get { return _fontColor; }
            set
            {
                _fontColor = value;
                _cachedHashCode = 0;
            }
        }

        public string FontName
        {
            get { return _fontName; }
            set
            {
                _fontName = value;
                _cachedHashCode = 0;
            }
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get { return _fontFamilyNumbering; }
            set
            {
                _fontFamilyNumbering = value;
                _cachedHashCode = 0;
            }
        }

        public XLFontCharSet FontCharSet
        {
            get { return _fontCharSet; }
            set
            {
                _fontCharSet = value;
                _cachedHashCode = 0;
            }
        }

        public bool Equals(XLFontKey other)
        {
            return
                Bold == other.Bold
             && Italic == other.Italic
             && Underline == other.Underline
             && Strikethrough == other.Strikethrough
             && VerticalAlignment == other.VerticalAlignment
             && Shadow == other.Shadow
             && FontSize == other.FontSize
             && FontColor == other.FontColor
             && FontFamilyNumbering == other.FontFamilyNumbering
             && FontCharSet == other.FontCharSet
             && string.Equals(FontName, other.FontName, StringComparison.InvariantCultureIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            if (obj is XLFontKey)
                return Equals((XLFontKey)obj);
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            if (_cachedHashCode != 0)
            {
                return _cachedHashCode;
            }

            var hashCode = 1158783753;
            hashCode = hashCode * -1521134295 + Bold.GetHashCode();
            hashCode = hashCode * -1521134295 + Italic.GetHashCode();
            hashCode = hashCode * -1521134295 + (int)Underline;
            hashCode = hashCode * -1521134295 + Strikethrough.GetHashCode();
            hashCode = hashCode * -1521134295 + (int)VerticalAlignment;
            hashCode = hashCode * -1521134295 + Shadow.GetHashCode();
            hashCode = hashCode * -1521134295 + FontSize.GetHashCode();
            hashCode = hashCode * -1521134295 + FontColor.GetHashCode();
            hashCode = hashCode * -1521134295 + StringComparer.InvariantCultureIgnoreCase.GetHashCode(FontName);
            hashCode = hashCode * -1521134295 + FontFamilyNumbering.GetHashCode();
            hashCode = hashCode * -1521134295 + FontCharSet.GetHashCode();

            if (hashCode == 0) hashCode = 1;
            _cachedHashCode = hashCode;

            return hashCode;
        }

        public override string ToString()
        {
            return $"{FontName} {FontSize}pt {FontColor} " +
                   (Bold ? "Bold" : "") + (Italic ? "Italic" : "") + (Strikethrough ? "Strikethrough" : "") +
                   (Underline == XLFontUnderlineValues.None ? "" : Underline.ToString()) +
                   $"{FontFamilyNumbering} {FontCharSet}";
        }

        public static bool operator ==(XLFontKey left, XLFontKey right) => left.Equals(right);

        public static bool operator !=(XLFontKey left, XLFontKey right) => !(left.Equals(right));
    }
}
