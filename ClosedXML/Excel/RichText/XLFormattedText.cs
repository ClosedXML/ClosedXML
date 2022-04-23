using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFormattedText<T> : IXLFormattedText<T>
    {
        private List<IXLRichString> _richTexts = new List<IXLRichString>();

        protected T Container;
        private readonly IXLFontBase _defaultFont;

        public XLFormattedText(IXLFontBase defaultFont)
        {
            Length = 0;
            _defaultFont = defaultFont;
        }

        public XLFormattedText(IXLFormattedText<T> defaultRichText, IXLFontBase defaultFont)
            : this(defaultFont)
        {
            foreach (var rt in defaultRichText)
                AddText(rt.Text, rt);
            if (defaultRichText.HasPhonetics)
            {
                _phonetics = new XLPhonetics(defaultRichText.Phonetics, defaultFont);
            }
        }

        public XLFormattedText(string text, IXLFontBase defaultFont)
            : this(defaultFont)
        {
            AddText(text);
        }

        public int Count
        { get { return _richTexts.Count; } }
        public int Length { get; private set; }

        public IXLRichString AddText(string text)
        {
            return AddText(text, _defaultFont);
        }

        public IXLRichString AddText(string text, IXLFontBase font)
        {
            var richText = new XLRichString(text, font, this);
            return AddText(richText);
        }

        public IXLRichString AddText(IXLRichString richText)
        {
            _richTexts.Add(richText);
            Length += richText.Text.Length;
            return richText;
        }

        public IXLRichString AddNewLine()
        {
            return AddText(XLConstants.NewLine);
        }

        public IXLFormattedText<T> ClearText()
        {
            _richTexts.Clear();
            Length = 0;
            return this;
        }

        public IXLFormattedText<T> ClearFont()
        {
            string text = Text;
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

        public IXLFormattedText<T> Substring(int index)
        {
            return Substring(index, Length - index);
        }

        public IXLFormattedText<T> Substring(int index, int length)
        {
            if (index + 1 > Length || (Length - index + 1) < length || length <= 0)
                throw new IndexOutOfRangeException("Index and length must refer to a location within the string.");

            List<IXLRichString> newRichTexts = new List<IXLRichString>();
            var retVal = new XLFormattedText<T>(_defaultFont);

            int lastPosition = 0;
            foreach (var rt in _richTexts)
            {
                if (lastPosition >= index + 1 + length) // We already have what we need
                {
                    newRichTexts.Add(rt);
                }
                else if (lastPosition + rt.Text.Length >= index + 1) // Eureka!
                {
                    int startIndex = index - lastPosition;

                    if (startIndex > 0)
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(0, startIndex), rt, this));
                    else if (startIndex < 0)
                        startIndex = 0;

                    int leftToTake = length - retVal.Length;
                    if (leftToTake > rt.Text.Length - startIndex)
                        leftToTake = rt.Text.Length - startIndex;

                    XLRichString newRt = new XLRichString(rt.Text.Substring(startIndex, leftToTake), rt, this);
                    newRichTexts.Add(newRt);
                    retVal.AddText(newRt);

                    if (startIndex + leftToTake < rt.Text.Length)
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(startIndex + leftToTake), rt, this));
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

        public bool Bold
        { set { _richTexts.ForEach(rt => rt.Bold = value); } }
        public bool Italic
        { set { _richTexts.ForEach(rt => rt.Italic = value); } }
        public XLFontUnderlineValues Underline
        { set { _richTexts.ForEach(rt => rt.Underline = value); } }
        public bool Strikethrough
        { set { _richTexts.ForEach(rt => rt.Strikethrough = value); } }
        public XLFontVerticalTextAlignmentValues VerticalAlignment
        { set { _richTexts.ForEach(rt => rt.VerticalAlignment = value); } }
        public bool Shadow
        { set { _richTexts.ForEach(rt => rt.Shadow = value); } }
        public double FontSize
        { set { _richTexts.ForEach(rt => rt.FontSize = value); } }
        public XLColor FontColor
        { set { _richTexts.ForEach(rt => rt.FontColor = value); } }
        public string FontName
        { set { _richTexts.ForEach(rt => rt.FontName = value); } }
        public XLFontFamilyNumberingValues FontFamilyNumbering
        { set { _richTexts.ForEach(rt => rt.FontFamilyNumbering = value); } }

        public IXLFormattedText<T> SetBold()
        { Bold = true; return this; }

        public IXLFormattedText<T> SetBold(bool value)
        { Bold = value; return this; }

        public IXLFormattedText<T> SetItalic()
        { Italic = true; return this; }

        public IXLFormattedText<T> SetItalic(bool value)
        { Italic = value; return this; }

        public IXLFormattedText<T> SetUnderline()
        { Underline = XLFontUnderlineValues.Single; return this; }

        public IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value)
        { Underline = value; return this; }

        public IXLFormattedText<T> SetStrikethrough()
        { Strikethrough = true; return this; }

        public IXLFormattedText<T> SetStrikethrough(bool value)
        { Strikethrough = value; return this; }

        public IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        { VerticalAlignment = value; return this; }

        public IXLFormattedText<T> SetShadow()
        { Shadow = true; return this; }

        public IXLFormattedText<T> SetShadow(bool value)
        { Shadow = value; return this; }

        public IXLFormattedText<T> SetFontSize(double value)
        { FontSize = value; return this; }

        public IXLFormattedText<T> SetFontColor(XLColor value)
        { FontColor = value; return this; }

        public IXLFormattedText<T> SetFontName(string value)
        { FontName = value; return this; }

        public IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        { FontFamilyNumbering = value; return this; }

        public bool Equals(IXLFormattedText<T> other)
        {
            int count = Count;
            if (count != other.Count)
                return false;

            for (int i = 0; i < count; i++)
            {
                if (_richTexts.ElementAt(i) != other.ElementAt(i))
                    return false;
            }

            return _phonetics == null || Phonetics.Equals(other.Phonetics);
        }

        public string Text
        { get { return ToString(); } }

        private IXLPhonetics _phonetics;

        public IXLPhonetics Phonetics
        {
            get { return _phonetics ?? (_phonetics = new XLPhonetics(_defaultFont)); }
        }

        public bool HasPhonetics
        { get { return _phonetics != null; } }
    }
}