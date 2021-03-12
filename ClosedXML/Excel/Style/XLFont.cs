using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFont : IXLFont
    {
        #region Static members

        public static IXLFontBase DefaultCommentFont
        {
            get
            {
                // MS Excel uses Tahoma 9 Swiss no matter what current style font
                var defaultCommentFont = new XLFont
                {
                    FontName = "Tahoma",
                    FontSize = 9,
                    FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss
                };

                return defaultCommentFont;
            }
        }

        internal static XLFontKey GenerateKey(IXLFontBase defaultFont)
        {
            if (defaultFont == null)
            {
                return XLFontValue.Default.Key;
            }
            else if (defaultFont is XLFont)
            {
                return (defaultFont as XLFont).Key;
            }
            else
            {
                return new XLFontKey
                {
                    Bold = defaultFont.Bold,
                    Italic = defaultFont.Italic,
                    Underline = defaultFont.Underline,
                    Strikethrough = defaultFont.Strikethrough,
                    VerticalAlignment = defaultFont.VerticalAlignment,
                    Shadow = defaultFont.Shadow,
                    FontSize = defaultFont.FontSize,
                    FontColor = defaultFont.FontColor.Key,
                    FontName = defaultFont.FontName,
                    FontFamilyNumbering = defaultFont.FontFamilyNumbering,
                    FontCharSet = defaultFont.FontCharSet
                };
            }
        }

        #endregion Static members

        private readonly XLStyle _style;

        private XLFontValue _value;

        internal XLFontKey Key
        {
            get { return _value.Key; }
            private set { _value = XLFontValue.FromKey(ref value); }
        }

        #region Constructors

        /// <summary>
        /// Create an instance of XLFont initializing it with the specified value.
        /// </summary>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLFont(XLStyle style, XLFontValue value)
        {
            _style = style ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLFont(XLStyle style, XLFontKey key) : this(style, XLFontValue.FromKey(ref key))
        {
        }

        public XLFont(XLStyle style = null, IXLFont d = null) : this(style, GenerateKey(d))
        {
        }

        #endregion Constructors

        private void Modify(Func<XLFontKey, XLFontKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var font = styleKey.Font;
                styleKey.Font = modification(font);
                return styleKey;
            });
        }

        #region IXLFont Members

        public Boolean Bold
        {
            get { return Key.Bold; }
            set
            {
                Modify(k => { k.Bold = value; return k; });
            }
        }

        public Boolean Italic
        {
            get { return Key.Italic; }
            set
            {
                Modify(k => { k.Italic = value; return k; });
            }
        }

        public XLFontUnderlineValues Underline
        {
            get { return Key.Underline; }
            set
            {
                Modify(k => { k.Underline = value; return k; });
            }
        }

        public Boolean Strikethrough
        {
            get { return Key.Strikethrough; }
            set
            {
                Modify(k => { k.Strikethrough = value; return k; });
            }
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get { return Key.VerticalAlignment; }
            set
            {
                Modify(k => { k.VerticalAlignment = value; return k; });
            }
        }

        public Boolean Shadow
        {
            get { return Key.Shadow; }
            set
            {
                Modify(k => { k.Shadow = value; return k; });
            }
        }

        public Double FontSize
        {
            get { return Key.FontSize; }
            set
            {
                Modify(k => { k.FontSize = value; return k; });
            }
        }

        public XLColor FontColor
        {
            get
            {
                var fontColorKey = Key.FontColor;
                return XLColor.FromKey(ref fontColorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");
                Modify(k => { k.FontColor = value.Key; return k; });
            }
        }

        public String FontName
        {
            get { return Key.FontName; }
            set
            {
                Modify(k => { k.FontName = value; return k; });
            }
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get { return Key.FontFamilyNumbering; }
            set
            {
                Modify(k => { k.FontFamilyNumbering = value; return k; });
            }
        }

        public XLFontCharSet FontCharSet
        {
            get { return Key.FontCharSet; }
            set
            {
                Modify(k => { k.FontCharSet = value; return k; });
            }
        }

        public IXLStyle SetBold()
        {
            Bold = true;
            return _style;
        }

        public IXLStyle SetBold(Boolean value)
        {
            Bold = value;
            return _style;
        }

        public IXLStyle SetItalic()
        {
            Italic = true;
            return _style;
        }

        public IXLStyle SetItalic(Boolean value)
        {
            Italic = value;
            return _style;
        }

        public IXLStyle SetUnderline()
        {
            Underline = XLFontUnderlineValues.Single;
            return _style;
        }

        public IXLStyle SetUnderline(XLFontUnderlineValues value)
        {
            Underline = value;
            return _style;
        }

        public IXLStyle SetStrikethrough()
        {
            Strikethrough = true;
            return _style;
        }

        public IXLStyle SetStrikethrough(Boolean value)
        {
            Strikethrough = value;
            return _style;
        }

        public IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        {
            VerticalAlignment = value;
            return _style;
        }

        public IXLStyle SetShadow()
        {
            Shadow = true;
            return _style;
        }

        public IXLStyle SetShadow(Boolean value)
        {
            Shadow = value;
            return _style;
        }

        public IXLStyle SetFontSize(Double value)
        {
            FontSize = value;
            return _style;
        }

        public IXLStyle SetFontColor(XLColor value)
        {
            FontColor = value;
            return _style;
        }

        public IXLStyle SetFontName(String value)
        {
            FontName = value;
            return _style;
        }

        public IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        {
            FontFamilyNumbering = value;
            return _style;
        }

        public IXLStyle SetFontCharSet(XLFontCharSet value)
        {
            FontCharSet = value;
            return _style;
        }

        #endregion IXLFont Members

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Bold.ToString());
            sb.Append("-");
            sb.Append(Italic.ToString());
            sb.Append("-");
            sb.Append(Underline.ToString());
            sb.Append("-");
            sb.Append(Strikethrough.ToString());
            sb.Append("-");
            sb.Append(VerticalAlignment.ToString());
            sb.Append("-");
            sb.Append(Shadow.ToString());
            sb.Append("-");
            sb.Append(FontSize.ToString());
            sb.Append("-");
            sb.Append(FontColor);
            sb.Append("-");
            sb.Append(FontName);
            sb.Append("-");
            sb.Append(FontFamilyNumbering.ToString());
            return sb.ToString();
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XLFont);
        }

        public Boolean Equals(IXLFont other)
        {
            var otherF = other as XLFont;
            if (otherF == null)
                return false;

            return Key == otherF.Key;
        }

        public override int GetHashCode()
        {
            var hashCode = 416600561;
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        #endregion Overridden
    }
}
