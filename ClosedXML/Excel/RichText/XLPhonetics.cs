using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPhonetics : IXLPhonetics
    {
        private readonly List<IXLPhonetic> _phonetics = new List<IXLPhonetic>();

        private readonly IXLFontBase _defaultFont;

        public XLPhonetics(IXLFontBase defaultFont)
        {
            _defaultFont = defaultFont;
            Type = XLPhoneticType.FullWidthKatakana;
            Alignment = XLPhoneticAlignment.Left;
            this.CopyFont(_defaultFont);
        }

        public XLPhonetics(IXLPhonetics defaultPhonetics, IXLFontBase defaultFont)
        {
            _defaultFont = defaultFont;
            Type = defaultPhonetics.Type;
            Alignment = defaultPhonetics.Alignment;

            this.CopyFont(defaultPhonetics);
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

        public IXLPhonetics Add(String text, Int32 start, Int32 end)
        {
            _phonetics.Add(new XLPhonetic(text, start, end));
            return this;
        }

        public IXLPhonetics ClearText()
        {
            _phonetics.Clear();
            return this;
        }

        public IXLPhonetics ClearFont()
        {
            this.CopyFont(_defaultFont);
            return this;
        }

        public Int32 Count { get { return _phonetics.Count; } }

        public XLPhoneticAlignment Alignment { get; set; }
        public XLPhoneticType Type { get; set; }

        public IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment) { Alignment = phoneticAlignment; return this; }

        public IXLPhonetics SetType(XLPhoneticType phoneticType) { Type = phoneticType; return this; }

        public IEnumerator<IXLPhonetic> GetEnumerator()
        {
            return _phonetics.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool Equals(IXLPhonetics other)
        {
            if (other == null)
                return false;

            Int32 phoneticsCount = _phonetics.Count;
            for (Int32 i = 0; i < phoneticsCount; i++)
            {
                if (!_phonetics[i].Equals(other.ElementAt(i)))
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
                && FontFamilyNumbering == other.FontFamilyNumbering;
        }
    }
}
