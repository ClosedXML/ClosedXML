using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRichText: IXLRichText
    {
        public XLRichText(String text)
            : this(text, XLWorkbook.DefaultStyle.Font)
        {
           
        }

        public XLRichText(String text, IXLFontBase defaultFont)
        {
            Text = text;

            Bold = defaultFont.Bold;
            Italic = defaultFont.Italic;
            Underline = defaultFont.Underline;
            Strikethrough = defaultFont.Strikethrough;
            VerticalAlignment = defaultFont.VerticalAlignment;
            Shadow = defaultFont.Shadow;
            FontSize = defaultFont.FontSize;
            FontColor = new XLColor(defaultFont.FontColor);
            FontName = defaultFont.FontName;
            FontFamilyNumbering = defaultFont.FontFamilyNumbering;
        }

        public String Text { get; private set; }

        public Boolean Bold { get; set; }
        public Boolean Italic { get; set; }
        public XLFontUnderlineValues Underline { get; set; }
        public Boolean Strikethrough { get; set; }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public Boolean Shadow { get; set; }
        public Double FontSize { get; set; }
        public IXLColor FontColor { get; set; }
        public String FontName { get; set; }
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

        public IXLRichText SetBold() { Bold = true; return this; }	public IXLRichText SetBold(Boolean value) { Bold = value; return this; }
        public IXLRichText SetItalic() { Italic = true; return this; }	public IXLRichText SetItalic(Boolean value) { Italic = value; return this; }
        public IXLRichText SetUnderline() { Underline = XLFontUnderlineValues.Single; return this; }	public IXLRichText SetUnderline(XLFontUnderlineValues value) { Underline = value; return this; }
        public IXLRichText SetStrikethrough() { Strikethrough = true; return this; }	public IXLRichText SetStrikethrough(Boolean value) { Strikethrough = value; return this; }
        public IXLRichText SetVerticalAlignment(XLFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }
        public IXLRichText SetShadow() { Shadow = true; return this; }	public IXLRichText SetShadow(Boolean value) { Shadow = value; return this; }
        public IXLRichText SetFontSize(Double value) { FontSize = value; return this; }
        public IXLRichText SetFontColor(IXLColor value) { FontColor = value; return this; }
        public IXLRichText SetFontName(String value) { FontName = value; return this; }
        public IXLRichText SetFontFamilyNumbering(XLFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

        public Boolean Equals(IXLRichText other)
        {
            return
                   this.Bold.Equals(other.Bold)
                && this.Italic.Equals(other.Italic)
                && this.Underline.Equals(other.Underline)
                && this.Strikethrough.Equals(other.Strikethrough)
                && this.VerticalAlignment.Equals(other.VerticalAlignment)
                && this.Shadow.Equals(other.Shadow)
                && this.FontSize.Equals(other.FontSize)
                && this.FontColor.Equals(other.FontColor)
                && this.FontName.Equals(other.FontName)
                && this.FontFamilyNumbering.Equals(other.FontFamilyNumbering)
                ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLRichText)obj);
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
