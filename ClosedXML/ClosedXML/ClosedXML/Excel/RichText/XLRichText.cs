using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRichText: IXLRichText, IEquatable<IXLRichText>
    {
        List<IXLRichString> richTexts = new List<IXLRichString>();

        IXLFontBase m_defaultFont;
        public XLRichText(IXLFontBase defaultFont)
        {
            m_defaultFont = defaultFont;
        }

        public XLRichText(String text, IXLFontBase defaultFont)
            :this(defaultFont)
        {
            AddText(text);
        }

        public Int32 Count { get { return richTexts.Count; } }
        private Int32 m_length = 0;
        public Int32 Length
        {
            get
            {
                return m_length;
            }
        }
        public IXLRichString AddText(String text)
        {
            return AddText(text, m_defaultFont);
        }
        public IXLRichString AddText(String text, IXLFontBase font)
        {
            var richText = new XLRichString(text, font);
            return AddText(richText);
        }

        public IXLRichString AddText(IXLRichString richText)
        {
            richTexts.Add(richText);
            m_length += richText.Text.Length;
            return richText;
        }
        public IXLRichText Clear()
        {
            richTexts.Clear();
            m_length = 0;
            return this;
        }

        public override string ToString()
        {
            var sb = new StringBuilder(richTexts.Count);
            richTexts.ForEach(rt => sb.Append(rt.Text));
            return sb.ToString();
        }

        public IXLRichText Substring(Int32 index)
        {
            return Substring(index, m_length - index);
        }
        public IXLRichText Substring(Int32 index, Int32 length)
        {
            if (index + 1 > m_length || (m_length - index + 1) < length || length <= 0)
                throw new IndexOutOfRangeException("Index and length must refer to a location within the string.");

            List<IXLRichString> newRichTexts = new List<IXLRichString>();
            XLRichText retVal = new XLRichText(m_defaultFont);

            Int32 lastPosition = 0;
            foreach (var rt in richTexts)
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
                    
                    XLRichString newRT = new XLRichString(rt.Text.Substring(startIndex, leftToTake), rt);
                    newRichTexts.Add(newRT);
                    retVal.AddText(newRT);

                    if (startIndex + leftToTake < rt.Text.Length)
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(startIndex + leftToTake), rt));
                }
                else // We haven't reached the desired position yet
                {
                    newRichTexts.Add(rt);
                }
                lastPosition += rt.Text.Length;
            }
            richTexts = newRichTexts;
            return retVal;
        }

        public IEnumerator<IXLRichString> GetEnumerator()
        {
            return richTexts.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Boolean Bold { set { richTexts.ForEach(rt => rt.Bold = value); } }
        public Boolean Italic { set { richTexts.ForEach(rt => rt.Italic = value); } }
        public XLFontUnderlineValues Underline { set { richTexts.ForEach(rt => rt.Underline = value); } }
        public Boolean Strikethrough { set { richTexts.ForEach(rt => rt.Strikethrough = value); } }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { set { richTexts.ForEach(rt => rt.VerticalAlignment = value); } }
        public Boolean Shadow { set { richTexts.ForEach(rt => rt.Shadow = value); } }
        public Double FontSize { set { richTexts.ForEach(rt => rt.FontSize = value); } }
        public IXLColor FontColor { set { richTexts.ForEach(rt => rt.FontColor = value); } }
        public String FontName { set { richTexts.ForEach(rt => rt.FontName = value); } }
        public XLFontFamilyNumberingValues FontFamilyNumbering { set { richTexts.ForEach(rt => rt.FontFamilyNumbering = value); } }

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
                if (richTexts.ElementAt(i) != other.ElementAt(i))
                    return false;
            }

            return true;
        }
    }
}
