using System;
using System.Linq;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLPhonetics : IXLPhonetics
    {
        private List<IXLPhonetic> phonetics = new List<IXLPhonetic>();

        IXLFontBase m_defaultFont;
        public XLPhonetics(IXLFontBase defaultFont)
        {
            m_defaultFont = defaultFont;
            Type = XLPhoneticType.FullWidthKatakana;
            Alignment = XLPhoneticAlignment.Left;
            ClearFont();
        }

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

        public IXLPhonetics SetBold() { Bold = true; return this; }	public IXLPhonetics SetBold(Boolean value) { Bold = value; return this; }
        public IXLPhonetics SetItalic() { Italic = true; return this; }	public IXLPhonetics SetItalic(Boolean value) { Italic = value; return this; }
        public IXLPhonetics SetUnderline() { Underline = XLFontUnderlineValues.Single; return this; }	public IXLPhonetics SetUnderline(XLFontUnderlineValues value) { Underline = value; return this; }
        public IXLPhonetics SetStrikethrough() { Strikethrough = true; return this; }	public IXLPhonetics SetStrikethrough(Boolean value) { Strikethrough = value; return this; }
        public IXLPhonetics SetVerticalAlignment(XLFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }
        public IXLPhonetics SetShadow() { Shadow = true; return this; }	public IXLPhonetics SetShadow(Boolean value) { Shadow = value; return this; }
        public IXLPhonetics SetFontSize(Double value) { FontSize = value; return this; }
        public IXLPhonetics SetFontColor(IXLColor value) { FontColor = value; return this; }
        public IXLPhonetics SetFontName(String value) { FontName = value; return this; }
        public IXLPhonetics SetFontFamilyNumbering(XLFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

        public IXLPhonetics Add(String text, Int32 start, Int32 end)
        {
            phonetics.Add(new XLPhonetic(text, start, end));
            return this;
        }
        public IXLPhonetics ClearText()
        {
            phonetics.Clear();
            return this;
        }
        public IXLPhonetics ClearFont()
        {
            Bold = m_defaultFont.Bold;
            Italic = m_defaultFont.Italic;
            Underline = m_defaultFont.Underline;
            Strikethrough = m_defaultFont.Strikethrough;
            VerticalAlignment = m_defaultFont.VerticalAlignment;
            Shadow = m_defaultFont.Shadow;
            FontSize = m_defaultFont.FontSize;
            FontColor = new XLColor(m_defaultFont.FontColor);
            FontName = m_defaultFont.FontName;
            FontFamilyNumbering = m_defaultFont.FontFamilyNumbering;
            return this;
        }
        public Int32 Count { get { return phonetics.Count; } }

        public XLPhoneticAlignment Alignment { get; set; }
        public XLPhoneticType Type { get; set; }

        public IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment) { Alignment = phoneticAlignment; return this; }
        public IXLPhonetics SetType(XLPhoneticType phoneticType) { Type = phoneticType; return this; }

        public IEnumerator<IXLPhonetic> GetEnumerator()
        {
            return phonetics.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool Equals(IXLPhonetics other)
        {
            if (other == null)
                return false;

            Int32 phoneticsCount = phonetics.Count;
            for (Int32 i = 0; i < phoneticsCount; i++)
            {
                if (!phonetics[i].Equals(other.ElementAt(i)))
                    return false;
            }

            return
                   Bold == other.Bold
                && Italic == other.Italic
                && Underline == other.Underline
                && Strikethrough == other.Strikethrough
                && VerticalAlignment == other.VerticalAlignment
                && Shadow == other.Shadow
                && FontSize == other.FontSize
                && FontColor.Equals(other.FontColor)
                && FontName == other.FontName
                && FontFamilyNumbering == other.FontFamilyNumbering
    ;
        }
    }
}
