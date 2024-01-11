using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPhonetics : IXLPhonetics
    {
        private readonly List<IXLPhonetic> _phonetics = new();
        private readonly XLFont _font;
        private readonly IXLFontBase _defaultFont;
        private readonly Action _onChange;
        private XLPhoneticAlignment _alignment;
        private XLPhoneticType _type;

        public XLPhonetics(IXLFontBase defaultFont, Action onChange)
        {
            _defaultFont = defaultFont;
            _font = new XLFont(defaultFont);
            _type = XLPhoneticType.FullWidthKatakana;
            _alignment = XLPhoneticAlignment.Left;
            _onChange = onChange;
        }

        public XLPhonetics(IXLPhonetics defaultPhonetics, IXLFontBase defaultFont, Action onChange)
        {
            _defaultFont = defaultFont;
            _font = new XLFont(defaultPhonetics);
            _type = defaultPhonetics.Type;
            _alignment = defaultPhonetics.Alignment;
            _onChange = onChange;
        }

        public Int32 Count => _phonetics.Count;

        public Boolean Bold
        {
            get => _font.Bold;
            set
            {
                _font.Bold = value;
                _onChange();
            }
        }

        public Boolean Italic
        {
            get => _font.Italic;
            set
            {
                _font.Italic = value;
                _onChange();
            }
        }

        public XLFontUnderlineValues Underline
        {
            get => _font.Underline;
            set
            {
                _font.Underline = value;
                _onChange();
            }
        }

        public Boolean Strikethrough
        {
            get => _font.Strikethrough;
            set
            {
                _font.Strikethrough = value;
                _onChange();
            }
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get => _font.VerticalAlignment;
            set
            {
                _font.VerticalAlignment = value;
                _onChange();
            }
        }

        public Boolean Shadow
        {
            get => _font.Shadow;
            set
            {
                _font.Shadow = value;
                _onChange();
            }
        }

        public Double FontSize
        {
            get => _font.FontSize;
            set
            {
                _font.FontSize = value;
                _onChange();
            }
        }

        public XLColor FontColor
        {
            get => _font.FontColor;
            set
            {
                _font.FontColor = value;
                _onChange();
            }
        }

        public String FontName
        {
            get => _font.FontName;
            set
            {
                _font.FontName = value;
                _onChange();
            }
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get => _font.FontFamilyNumbering;
            set
            {
                _font.FontFamilyNumbering = value;
                _onChange();
            }
        }

        public XLFontCharSet FontCharSet
        {
            get => _font.FontCharSet;
            set
            {
                _font.FontCharSet = value;
                _onChange();
            }
        }

        public XLFontScheme FontScheme
        {
            get => _font.FontScheme;
            set
            {
                _font.FontScheme = value;
                _onChange();
            }
        }

        public XLPhoneticAlignment Alignment
        {
            get => _alignment;
            set
            {
                _alignment = value;
                _onChange();
            }
        }

        public XLPhoneticType Type
        {
            get => _type;
            set
            {
                _type = value;
                _onChange();
            }
        }

        public IXLPhonetics SetBold() { Bold = true; return this; }

        public IXLPhonetics SetBold(Boolean value) { Bold = value; return this; }

        public IXLPhonetics SetItalic() { Italic = true; return this; }

        public IXLPhonetics SetItalic(Boolean value) { Italic = value; return this; }

        public IXLPhonetics SetUnderline() { Underline = XLFontUnderlineValues.Single; return this; }

        public IXLPhonetics SetUnderline(XLFontUnderlineValues value) { Underline = value; return this; }

        public IXLPhonetics SetStrikethrough() { Strikethrough = true; return this; }

        public IXLPhonetics SetStrikethrough(Boolean value) { Strikethrough = value; return this; }

        public IXLPhonetics SetVerticalAlignment(XLFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }

        public IXLPhonetics SetShadow() { Shadow = true; return this; }

        public IXLPhonetics SetShadow(Boolean value) { Shadow = value; return this; }

        public IXLPhonetics SetFontSize(Double value) { FontSize = value; return this; }

        public IXLPhonetics SetFontColor(XLColor value) { FontColor = value; return this; }

        public IXLPhonetics SetFontName(String value) { FontName = value; return this; }

        public IXLPhonetics SetFontFamilyNumbering(XLFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

        public IXLPhonetics SetFontCharSet(XLFontCharSet value) { FontCharSet = value; return this; }

        public IXLPhonetics SetFontScheme(XLFontScheme value) { FontScheme = value; return this; }

        public IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment) { Alignment = phoneticAlignment; return this; }

        public IXLPhonetics SetType(XLPhoneticType phoneticType) { Type = phoneticType; return this; }

        public IXLPhonetics Add(String text, Int32 start, Int32 end)
        {
            _phonetics.Add(new XLPhonetic(text, start, end));
            _onChange();
            return this;
        }

        public IXLPhonetics ClearText()
        {
            _phonetics.Clear();
            _onChange();
            return this;
        }

        public IXLPhonetics ClearFont()
        {
            this.CopyFont(_defaultFont);
            _onChange();
            return this;
        }

        public IEnumerator<IXLPhonetic> GetEnumerator()
        {
            return _phonetics.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool Equals(IXLPhonetics? other) => Equals(other as XLPhonetics);

        public bool Equals(XLPhonetics? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            if (!_phonetics.SequenceEqual(other._phonetics))
                return false;

            return
                _font.Key.Equals(other._font.Key) &&
                Type == other.Type &&
                Alignment == other.Alignment;
        }
    }
}
