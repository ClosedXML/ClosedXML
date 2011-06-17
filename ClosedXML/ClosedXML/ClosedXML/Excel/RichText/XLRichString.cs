using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.RichText
{
    internal class XLRichString: IXLRichString
    {
        List<IXLRichText> richTexts = new List<IXLRichText>();
        public IXLRichText AddText(String text)
        {
            var richText = new XLRichText(text);
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
    }
}
