using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRichString: IXLRichString
    {
        List<IXLRichText> richTexts = new List<IXLRichText>();

        IXLFontBase m_defaultFont;
        public XLRichString(IXLFontBase defaultFont)
        {
            m_defaultFont = defaultFont;
        }

        public XLRichString(String text, IXLFontBase defaultFont)
            :this(defaultFont)
        {
            AddText(text);
        }

        public Int32 Count { get { return richTexts.Count; } }
        public IXLRichText AddText(String text)
        {
            var richText = new XLRichText(text, m_defaultFont);
            richTexts.Add(richText);
            return richText;
        }
        public IXLRichString Clear()
        {
            richTexts.Clear();
            return this;
        }

        public override string ToString()
        {
            var sb = new StringBuilder(richTexts.Count);
            richTexts.ForEach(rt => sb.Append(rt.Text));
            return sb.ToString();
        }

        public IXLRichText Characters(Int32 index, Int32 length)
        {
            List<IXLRichText> newRichTexts = new List<IXLRichText>();
            Int32 runningLength = 0;
            foreach (var rt in richTexts)
            {
                if (runningLength + rt.Text.Length > index + 1)
                {
                    Int32 startIndex = index - runningLength + 1;
                    if (startIndex > 0)
                    {
                        var newRT = new XLRichText(rt.Text.Substring(0, startIndex + 1), rt);
                        newRichTexts.Add(newRT);
                    }

                    if (rt.Text.Length - startIndex + 1 >= length)
                    {
                        var newRT = new XLRichText(rt.Text.Substring(startIndex, length), rt);
                        newRichTexts.Add(newRT);

                        if (rt.Text.Length > startIndex + length + 1)
                        {
                            newRichTexts.Add(new XLRichText(rt.Text.Substring(startIndex + length), rt));
                        }
                    }

                    
                }
                else
                {
                    newRichTexts.Add(rt);
                    runningLength += rt.Text.Length;
                }
            }
            richTexts = newRichTexts;
            throw new NotImplementedException();
        }

        public IEnumerator<IXLRichText> GetEnumerator()
        {
            return richTexts.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
