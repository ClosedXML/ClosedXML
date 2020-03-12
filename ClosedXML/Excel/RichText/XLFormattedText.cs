using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLFormattedText<T>: IXLFormattedText<T>
    {
        List<IXLRichString> _richTexts = new List<IXLRichString>();

        protected T Container;
        readonly IXLFontBase _defaultFont;

        public XLFormattedText(IXLFontBase defaultFont)
        {
            Length = 0;
            _defaultFont = defaultFont;
            _isDefault = true;
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

            _isDefault = _richTexts.Count <= 1;
        }

        public XLFormattedText(String text, IXLFontBase defaultFont)
            :this(defaultFont)
        {
            AddText(text);
            _isDefault = true;
        }

        private bool _isDefault;

        /// <summary>
        /// Flag showing that the formatted text was constructed from a simple string and has not change since then.
        /// When IsDefault is true a rich text can be converted to an ordinary text without information loss.
        /// </summary>
        internal bool IsDefault => _isDefault && !HasPhonetics;

        public Int32 Count { get { return _richTexts.Count; } }
        public int Length { get; private set; }

        public IXLRichString AddText(String text)
        {
            return AddText(text, _defaultFont);
        }
        public IXLRichString AddText(String text, IXLFontBase font)
        {
            var richText = new XLRichString(text, font, this);
            return AddText(richText);
        }

        public IXLRichString AddText(IXLRichString richText)
        {
            _richTexts.Add(richText);
            Length += richText.Text.Length;
            _isDefault = false;
            return richText;
        }

        public IXLRichString AddNewLine()
        {
            return AddText(Environment.NewLine);
        }

        public IXLFormattedText<T> ClearText()
        {
            _richTexts.Clear();
            Length = 0;
            _isDefault = true;
            return this;
        }
        public IXLFormattedText<T> ClearFont()
        {
            String text = Text;
            ClearText();
            AddText(text);
            _isDefault = true;
            return this;
        }

        public override string ToString()
        {
            var sb = new StringBuilder(_richTexts.Count);
            _richTexts.ForEach(rt => sb.Append(rt.Text));
            return sb.ToString();
        }

        public IXLFormattedText<T> Substring(Int32 index)
        {
            return Substring(index, Length - index);
        }
        public IXLFormattedText<T> Substring(Int32 index, Int32 length)
        {
            if (index + 1 > Length || (Length - index + 1) < length || length <= 0)
                throw new IndexOutOfRangeException("Index and length must refer to a location within the string.");

            List<IXLRichString> newRichTexts = new List<IXLRichString>();
            var retVal = new XLFormattedText<T>(_defaultFont);

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
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(0, startIndex), rt, this));
                    else if (startIndex < 0)
                        startIndex = 0;

                    Int32 leftToTake = length - retVal.Length;
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
            _isDefault = false;
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

        public Boolean Bold
        {
            set
            {
                _richTexts.ForEach(rt => rt.Bold = value);
                _isDefault = false;
            }
        }

        public Boolean Italic
        {
            set
            {
                _richTexts.ForEach(rt => rt.Italic = value);
                _isDefault = false;
            }
        }

        public XLFontUnderlineValues Underline
        {
            set
            {
                _richTexts.ForEach(rt => rt.Underline = value);
                _isDefault = false;
            }
        }

        public Boolean Strikethrough
        {
            set
            {
                _richTexts.ForEach(rt => rt.Strikethrough = value);
                _isDefault = false;
            }
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            set
            {
                _richTexts.ForEach(rt => rt.VerticalAlignment = value);
                _isDefault = false;
            }
        }

        public Boolean Shadow
        {
            set
            {
                _richTexts.ForEach(rt => rt.Shadow = value);
                _isDefault = false;
            }
        }

        public Double FontSize
        {
            set
            {
                _richTexts.ForEach(rt => rt.FontSize = value);
                _isDefault = false;
            }
        }

        public XLColor FontColor
        {
            set
            {
                _richTexts.ForEach(rt => rt.FontColor = value);
                _isDefault = false;
            }
        }

        public String FontName
        {
            set
            {
                _richTexts.ForEach(rt => rt.FontName = value);
                _isDefault = false;
            }
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            set
            {
                _richTexts.ForEach(rt => rt.FontFamilyNumbering = value);
                _isDefault = false;
            }
        }

        public IXLFormattedText<T> SetBold() { Bold = true; return this; }	public IXLFormattedText<T> SetBold(Boolean value) { Bold = value; return this; }
        public IXLFormattedText<T> SetItalic() { Italic = true; return this; }	public IXLFormattedText<T> SetItalic(Boolean value) { Italic = value; return this; }
        public IXLFormattedText<T> SetUnderline() { Underline = XLFontUnderlineValues.Single; return this; }	public IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value) { Underline = value; return this; }
        public IXLFormattedText<T> SetStrikethrough() { Strikethrough = true; return this; }	public IXLFormattedText<T> SetStrikethrough(Boolean value) { Strikethrough = value; return this; }
        public IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }
        public IXLFormattedText<T> SetShadow() { Shadow = true; return this; }	public IXLFormattedText<T> SetShadow(Boolean value) { Shadow = value; return this; }
        public IXLFormattedText<T> SetFontSize(Double value) { FontSize = value; return this; }
        public IXLFormattedText<T> SetFontColor(XLColor value) { FontColor = value; return this; }
        public IXLFormattedText<T> SetFontName(String value) { FontName = value; return this; }
        public IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

        public bool Equals(IXLFormattedText<T> other)
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
