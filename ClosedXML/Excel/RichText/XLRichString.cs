using ClosedXML.Extensions;
using ClosedXML.Extensions;
using System.Diagnostics;

namespace ClosedXML.Excel.RichText
{
    [DebuggerDisplay("{Text}")]
    internal class XLRichString : IXLRichString
    {
        private readonly IXLWithRichString _withRichString;

        public XLRichString(string text, IXLFontBase font, IXLWithRichString withRichString)
        {
            Text = text;
            this.CopyFont(font);
            _withRichString = withRichString;
        }

        public string Text { get; set; }

        public IXLRichString AddText(string text)
        {
            return _withRichString.AddText(text);
        }

        public IXLRichString AddNewLine()
        {
            return AddText(XLConstants.NewLine);
        }

        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public XLFontUnderlineValues Underline { get; set; }
        public bool Strikethrough { get; set; }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public bool Shadow { get; set; }
        public double FontSize { get; set; }
        public XLColor FontColor { get; set; }
        public string FontName { get; set; }
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }
        public XLFontCharSet FontCharSet { get; set; }

        public IXLRichString SetBold()
        {
            Bold = true; return this;
        }

        public IXLRichString SetBold(bool value)
        {
            Bold = value; return this;
        }

        public IXLRichString SetItalic()
        {
            Italic = true; return this;
        }

        public IXLRichString SetItalic(bool value)
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

        public IXLRichString SetStrikethrough(bool value)
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

        public IXLRichString SetShadow(bool value)
        {
            Shadow = value; return this;
        }

        public IXLRichString SetFontSize(double value)
        {
            FontSize = value; return this;
        }

        public IXLRichString SetFontColor(XLColor value)
        {
            FontColor = value; return this;
        }

        public IXLRichString SetFontName(string value)
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

        public bool Equals(IXLRichString other)
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
                ^ (int)Underline
                ^ Strikethrough.GetHashCode()
                ^ (int)VerticalAlignment
                ^ Shadow.GetHashCode()
                ^ FontSize.GetHashCode()
                ^ FontColor.GetHashCode()
                ^ FontName.GetHashCode()
                ^ (int)FontFamilyNumbering;
        }
    }
}