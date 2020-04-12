using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Text}")]
    internal class XLRichString : IXLRichString
    {
        private IXLWithRichString _withRichString;

        public XLRichString(String text, IXLFontBase font, IXLWithRichString withRichString)
        {
            Text = text;
            this.CopyFont(font);
            _withRichString = withRichString;
        }

        public String Text { get; set; }

        public IXLRichString AddText(String text)
        {
            return _withRichString.AddText(text);
        }

        public IXLRichString AddNewLine()
        {
            return AddText(Environment.NewLine);
        }

        public Boolean Bold { get; set; }
        public Boolean Italic { get; set; }
        public XLFontUnderlineValues Underline { get; set; }
        public Boolean Strikethrough { get; set; }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public Boolean Shadow { get; set; }
        public Double FontSize { get; set; }
        public XLColor FontColor { get; set; }
        public String FontName { get; set; }
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }
        public XLFontCharSet FontCharSet { get; set; }

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

        public Boolean Equals(IXLRichString other)
        {
            return
                    Text == other.Text
                && Bold.Equals(other.Bold)
                && Italic.Equals(other.Italic)
                && Underline.Equals(other.Underline)
                && Strikethrough.Equals(other.Strikethrough)
                && VerticalAlignment.Equals(other.VerticalAlignment)
                && Shadow.Equals(other.Shadow)
                && FontSize.Equals(other.FontSize)
                && FontColor.Equals(other.FontColor)
                && FontName.Equals(other.FontName)
                && FontFamilyNumbering.Equals(other.FontFamilyNumbering)
                ;
        }

        public override bool Equals(object obj)
        {
            return Equals((XLRichString)obj);
        }

        public override int GetHashCode()
        {
            return Text.GetHashCode()
                ^ Bold.GetHashCode()
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
