using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Text}")]
    internal class XLRichString : IXLRichString
    {
        private readonly IXLWithRichString _withRichString;
        private readonly XLFont _font;
        private readonly Action _onChange;
        private string _text;

        public XLRichString(String text, IXLFontBase font, IXLWithRichString withRichString, Action? onChange)
        {
            _text = text;
            _font = new XLFont(font);
            _withRichString = withRichString;
            _onChange = onChange ?? (() => { });
        }

        public String Text
        {
            get => _text;
            set
            {
                _text = value;
                _onChange();
            }
        }

        public IXLRichString AddText(String text)
        {
            return _withRichString.AddText(text);
        }

        public IXLRichString AddNewLine()
        {
            return AddText(Environment.NewLine);
        }

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

        public IXLRichString SetBold()
        {
            Bold = true; return this;
        }

        public IXLRichString SetBold(Boolean value)
        {
            Bold = value; return this;
        }

        public IXLRichString SetItalic()
        {
            Italic = true; return this;
        }

        public IXLRichString SetItalic(Boolean value)
        {
            Italic = value; return this;
        }

        public IXLRichString SetUnderline()
        {
            Underline = XLFontUnderlineValues.Single; return this;
        }

        public IXLRichString SetUnderline(XLFontUnderlineValues value)
        {
            Underline = value; return this;
        }

        public IXLRichString SetStrikethrough()
        {
            Strikethrough = true; return this;
        }

        public IXLRichString SetStrikethrough(Boolean value)
        {
            Strikethrough = value; return this;
        }

        public IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        {
            VerticalAlignment = value; return this;
        }

        public IXLRichString SetShadow()
        {
            Shadow = true; return this;
        }

        public IXLRichString SetShadow(Boolean value)
        {
            Shadow = value; return this;
        }

        public IXLRichString SetFontSize(Double value)
        {
            FontSize = value; return this;
        }

        public IXLRichString SetFontColor(XLColor value)
        {
            FontColor = value; return this;
        }

        public IXLRichString SetFontName(String value)
        {
            FontName = value; return this;
        }

        public IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        {
            FontFamilyNumbering = value; return this;
        }

        public IXLRichString SetFontCharSet(XLFontCharSet value)
        {
            FontCharSet = value; return this;
        }

        public IXLRichString SetFontScheme(XLFontScheme value)
        {
            FontScheme = value; return this;
        }

        public override bool Equals(object? obj) => Equals(obj as XLRichString);

        public Boolean Equals(IXLRichString? other) => Equals(other as XLRichString);

        public Boolean Equals(XLRichString? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            return Text == other.Text && _font.Key.Equals(other._font.Key);
        }

        public override int GetHashCode()
        {
            // Since all properties of type are mutable, can't have different hashcode for any instance.
            // Don't ever use this class in a dictionary, e.g. SST.
            return 4; // Chosen by fair dice roll. Guaranteed to be random.
        }
    }
}
