using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFont : IXLFont
    {
        private readonly IXLStylized _container;
        private Boolean _bold;
        private XLColor _fontColor;
        private XLFontFamilyNumberingValues _fontFamilyNumbering;
        private String _fontName;
        private Double _fontSize;
        private Boolean _italic;
        private Boolean _shadow;
        private Boolean _strikethrough;
        private XLFontUnderlineValues _underline;
        private XLFontVerticalTextAlignmentValues _verticalAlignment;

        public XLFont()
            : this(null, XLWorkbook.DefaultStyle.Font)
        {
        }

        public XLFont(IXLStylized container, IXLFontBase defaultFont, Boolean useDefaultModify = true)
        {
            _container = container;
            if (defaultFont == null) return;

            _bold = defaultFont.Bold;
            _italic = defaultFont.Italic;
            _underline = defaultFont.Underline;
            _strikethrough = defaultFont.Strikethrough;
            _verticalAlignment = defaultFont.VerticalAlignment;
            _shadow = defaultFont.Shadow;
            _fontSize = defaultFont.FontSize;
            _fontColor = defaultFont.FontColor;
            _fontName = defaultFont.FontName;
            _fontFamilyNumbering = defaultFont.FontFamilyNumbering;

            if (useDefaultModify)
            {
                var d = defaultFont as XLFont;
                if (d == null) return;
                BoldModified = d.BoldModified;
                ItalicModified = d.ItalicModified;
                UnderlineModified = d.UnderlineModified;
                StrikethroughModified = d.StrikethroughModified;
                VerticalAlignmentModified = d.VerticalAlignmentModified;
                ShadowModified = d.ShadowModified;
                FontSizeModified = d.FontSizeModified;
                FontColorModified = d.FontColorModified;
                FontNameModified = d.FontNameModified;
                FontFamilyNumberingModified = d.FontFamilyNumberingModified;
            }
        }

        #region IXLFont Members

        public Boolean BoldModified { get; set; }
        public Boolean Bold
        {
            get { return _bold; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.Bold = value);
                else
                {
                    _bold = value;
                    BoldModified = true;
                }
            }
        }

        public Boolean ItalicModified { get; set; }
        public Boolean Italic
        {
            get { return _italic; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.Italic = value);
                else
                {
                    _italic = value;
                    ItalicModified = true;
                }
            }
        }

        public Boolean UnderlineModified { get; set; }
        public XLFontUnderlineValues Underline
        {
            get { return _underline; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.Underline = value);
                else
                {
                    _underline = value;
                    UnderlineModified = true;
                }
                    
            }
        }

        public Boolean StrikethroughModified { get; set; }
        public Boolean Strikethrough
        {
            get { return _strikethrough; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.Strikethrough = value);
                else
                {
                    _strikethrough = value;
                    StrikethroughModified = true;
                }
            }
        }

        public Boolean VerticalAlignmentModified { get; set; }
        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get { return _verticalAlignment; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.VerticalAlignment = value);
                else
                {
                    _verticalAlignment = value;
                    VerticalAlignmentModified = true;
                }
            }
        }

        public Boolean ShadowModified { get; set; }
        public Boolean Shadow
        {
            get { return _shadow; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.Shadow = value);
                else
                {
                    _shadow = value;
                    ShadowModified = true;
                }
            }
        }

        public Boolean FontSizeModified { get; set; }
        public Double FontSize
        {
            get { return _fontSize; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.FontSize = value);
                else
                {
                    _fontSize = value;
                    FontSizeModified = true;
                }
            }
        }

        private Boolean _fontColorModified;
        public Boolean FontColorModified
        {
            get { return _fontColorModified; }
            set
            {
                _fontColorModified = value;
            }
        }
        public XLColor FontColor
        {
            get { return _fontColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.FontColor = value);
                else
                {
                    _fontColor = value;
                    FontColorModified = true;
                }
            }
        }

        public Boolean FontNameModified { get; set; }
        public String FontName
        {
            get { return _fontName; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.FontName = value);
                else
                {
                    _fontName = value;
                    FontNameModified = true;
                }
            }
        }

        public Boolean FontFamilyNumberingModified { get; set; }
        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get { return _fontFamilyNumbering; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Font.FontFamilyNumbering = value);
                else
                {
                    _fontFamilyNumbering = value;
                    FontFamilyNumberingModified = true;
                }
            }
        }

        public IXLStyle SetBold()
        {
            Bold = true;
            return _container.Style;
        }

        public IXLStyle SetBold(Boolean value)
        {
            Bold = value;
            return _container.Style;
        }

        public IXLStyle SetItalic()
        {
            Italic = true;
            return _container.Style;
        }

        public IXLStyle SetItalic(Boolean value)
        {
            Italic = value;
            return _container.Style;
        }

        public IXLStyle SetUnderline()
        {
            Underline = XLFontUnderlineValues.Single;
            return _container.Style;
        }

        public IXLStyle SetUnderline(XLFontUnderlineValues value)
        {
            Underline = value;
            return _container.Style;
        }

        public IXLStyle SetStrikethrough()
        {
            Strikethrough = true;
            return _container.Style;
        }

        public IXLStyle SetStrikethrough(Boolean value)
        {
            Strikethrough = value;
            return _container.Style;
        }

        public IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        {
            VerticalAlignment = value;
            return _container.Style;
        }

        public IXLStyle SetShadow()
        {
            Shadow = true;
            return _container.Style;
        }

        public IXLStyle SetShadow(Boolean value)
        {
            Shadow = value;
            return _container.Style;
        }

        public IXLStyle SetFontSize(Double value)
        {
            FontSize = value;
            return _container.Style;
        }

        public IXLStyle SetFontColor(XLColor value)
        {
            FontColor = value;
            return _container.Style;
        }

        public IXLStyle SetFontName(String value)
        {
            FontName = value;
            return _container.Style;
        }

        public IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        {
            FontFamilyNumbering = value;
            return _container.Style;
        }

        public Boolean Equals(IXLFont other)
        {
            var otherF = other as XLFont;
            if (otherF == null)
                return false;

            return
                _bold == otherF._bold
                && _italic == otherF._italic
                && _underline == otherF._underline
                && _strikethrough == otherF._strikethrough
                && _verticalAlignment == otherF._verticalAlignment
                && _shadow == otherF._shadow
                && _fontSize == otherF._fontSize
                && _fontColor.Equals(otherF._fontColor)
                && _fontName == otherF._fontName
                && _fontFamilyNumbering == otherF._fontFamilyNumbering
                ;
        }

        #endregion

        private void SetStyleChanged()
        {
            if (_container != null) _container.StyleChanged = true;
        }

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
            return Equals((XLFont)obj);
        }

        public override int GetHashCode()
        {
            return Bold.GetHashCode()
                   ^ Italic.GetHashCode()
                   ^ (Int32)Underline
                   ^ Strikethrough.GetHashCode()
                   ^ (Int32)VerticalAlignment
                   ^ Shadow.GetHashCode()
                   ^ FontSize.GetHashCode()
                   ^ FontColor.GetHashCode()
                   ^ FontName.GetHashCode()
                   ^ (Int32)FontFamilyNumbering;
        }
    }
}