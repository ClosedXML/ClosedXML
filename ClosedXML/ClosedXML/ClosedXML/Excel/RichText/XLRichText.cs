using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRichText: IXLRichText
    {
        List<IXLRichString> _richTexts = new List<IXLRichString>();

        readonly IXLFontBase _defaultFont;
        public XLRichText(IXLFontBase defaultFont)
        {
            _defaultFont = defaultFont;
        }

        public XLRichText(IXLRichText defaultRichText, IXLFontBase defaultFont)
        {
            _defaultFont = defaultFont;
            foreach (var rt in defaultRichText)
                AddText(rt.Text, rt);
            if (defaultRichText.HasPhonetics)
            {
                _phonetics = new XLPhonetics(defaultRichText.Phonetics, defaultFont);
            }
        }

        public XLRichText(String text, IXLFontBase defaultFont)
            :this(defaultFont)
        {
            AddText(text);
        }

        public Int32 Count { get { return _richTexts.Count; } }
        private Int32 _length = 0;
        public Int32 Length
        {
            get
            {
                return _length;
            }
        }
        public IXLRichString AddText(String text)
        {
            return AddText(text, _defaultFont);
        }
        public IXLRichString AddText(String text, IXLFontBase font)
        {
            var richText = new XLRichString(text, font);
            return AddText(richText);
        }

        public IXLRichString AddText(IXLRichString richText)
        {
            _richTexts.Add(richText);
            _length += richText.Text.Length;
            return richText;
        }
        public IXLRichText ClearText()
        {
            _richTexts.Clear();
            _length = 0;
            return this;
        }
        public IXLRichText ClearFont()
        {
            String text = Text;
            ClearText();
            AddText(text);
            return this;
        }

        public override string ToString()
        {
            var sb = new StringBuilder(_richTexts.Count);
            _richTexts.ForEach(rt => sb.Append(rt.Text));
            return sb.ToString();
        }

        public IXLRichText Substring(Int32 index)
        {
            return Substring(index, _length - index);
        }
        public IXLRichText Substring(Int32 index, Int32 length)
        {
            if (index + 1 > _length || (_length - index + 1) < length || length <= 0)
                throw new IndexOutOfRangeException("Index and length must refer to a location within the string.");

            List<IXLRichString> newRichTexts = new List<IXLRichString>();
            XLRichText retVal = new XLRichText(_defaultFont);

            Int32 lastPosition = 0;
            foreach (var rt in _richTexts)
            {
                if (lastPosition >= index + 1 + length) // We already have what we need
                {
                    newRichTexts.Add(rt);
                }
                else if (lastPosition + rt.Text.Length >= index + 1) // Eureka!
                {
                    Int32 startIndex = index - lastPosition;

                    if (startIndex > 0)
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(0, startIndex), rt));
                    else if (startIndex < 0)
                        startIndex = 0;

                    Int32 leftToTake = length - retVal.Length;
                    if (leftToTake > rt.Text.Length - startIndex)
                        leftToTake = rt.Text.Length - startIndex;
                    
                    XLRichString newRt = new XLRichString(rt.Text.Substring(startIndex, leftToTake), rt);
                    newRichTexts.Add(newRt);
                    retVal.AddText(newRt);

                    if (startIndex + leftToTake < rt.Text.Length)
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(startIndex + leftToTake), rt));
                }
                else // We haven't reached the desired position yet
                {
                    newRichTexts.Add(rt);
                }
                lastPosition += rt.Text.Length;
            }
            _richTexts = newRichTexts;
            return retVal;
        }

        public IEnumerator<IXLRichString> GetEnumerator()
        {
            return _richTexts.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Boolean Bold { set { _richTexts.ForEach(rt => rt.Bold = value); } }
        public Boolean Italic { set { _richTexts.ForEach(rt => rt.Italic = value); } }
        public XLFontUnderlineValues Underline { set { _richTexts.ForEach(rt => rt.Underline = value); } }
        public Boolean Strikethrough { set { _richTexts.ForEach(rt => rt.Strikethrough = value); } }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { set { _richTexts.ForEach(rt => rt.VerticalAlignment = value); } }
        public Boolean Shadow { set { _richTexts.ForEach(rt => rt.Shadow = value); } }
        public Double FontSize { set { _richTexts.ForEach(rt => rt.FontSize = value); } }
        public IXLColor FontColor { set { _richTexts.ForEach(rt => rt.FontColor = value); } }
        public String FontName { set { _richTexts.ForEach(rt => rt.FontName = value); } }
        public XLFontFamilyNumberingValues FontFamilyNumbering { set { _richTexts.ForEach(rt => rt.FontFamilyNumbering = value); } }

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

        public bool Equals(IXLRichText other)
        {
            Int32 count = Count;
            if (count != other.Count)
                return false;

            for (Int32 i = 0; i < count; i++)
            {
                if (_richTexts.ElementAt(i) != other.ElementAt(i))
                    return false;
            }

            return _phonetics == null || Phonetics.Equals(other.Phonetics);
        }

        public String Text { get { return ToString(); } }

        private IXLPhonetics _phonetics;
        public IXLPhonetics Phonetics 
        {
            get { return _phonetics ?? (_phonetics = new XLPhonetics(_defaultFont)); }
        }

        public Boolean HasPhonetics { get { return _phonetics != null; } }
    }
}
