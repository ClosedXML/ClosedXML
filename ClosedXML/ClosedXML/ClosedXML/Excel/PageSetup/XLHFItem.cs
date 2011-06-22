using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLHFItem : IXLHFItem
    {
        public XLHFItem()
        { }
        public XLHFItem(XLHFItem defaultHFItem)
        {
            defaultHFItem.texts.ForEach(kp => texts.Add(kp.Key, kp.Value));
        }
        private Dictionary<XLHFOccurrence, List<XLHFText>> texts = new Dictionary<XLHFOccurrence, List<XLHFText>>();
        public String GetText(XLHFOccurrence occurrence)
        {
            var sb = new StringBuilder();
            if(texts.ContainsKey(occurrence))
            {
                foreach (var hfText in texts[occurrence])
                    sb.Append(hfText.HFText);
            }

            return sb.ToString();
        }

        public IXLRichString AddText(String text)
        {
            return AddText(text, XLHFOccurrence.AllPages);
        }
        public IXLRichString AddText(XLHFPredefinedText predefinedText)
        {
            return AddText(predefinedText, XLHFOccurrence.AllPages);
        }

        public IXLRichString AddText(String text, XLHFOccurrence occurrence)
        {
            IXLRichString richText = new XLRichString(text, XLWorkbook.DefaultStyle.Font);

            var hfText = new XLHFText(richText);
            if (occurrence == XLHFOccurrence.AllPages)
            {
                AddTextToOccurrence(hfText, XLHFOccurrence.EvenPages);
                AddTextToOccurrence(hfText, XLHFOccurrence.FirstPage);
                AddTextToOccurrence(hfText, XLHFOccurrence.OddPages);
            }
            else
            {
                AddTextToOccurrence(hfText, occurrence);
            }

            return richText;
        }

        private void AddTextToOccurrence(XLHFText hfText, XLHFOccurrence occurrence)
        {
            if (texts.ContainsKey(occurrence))
                texts[occurrence].Add(hfText);
            else
                texts.Add(occurrence, new List<XLHFText>() { hfText });
        }

        public IXLRichString AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence)
        {
            String hfText;
            switch (predefinedText)
            {
                case XLHFPredefinedText.PageNumber: hfText = "&P"; break;
                case XLHFPredefinedText.NumberOfPages : hfText = "&N"; break;
                case XLHFPredefinedText.Date : hfText = "&D"; break;
                case XLHFPredefinedText.Time : hfText = "&T"; break;
                case XLHFPredefinedText.Path : hfText = "&Z"; break;
                case XLHFPredefinedText.File : hfText = "&F"; break;
                case XLHFPredefinedText.SheetName : hfText = "&A"; break;
                case XLHFPredefinedText.FullPath: hfText = "&Z&F"; break;
                default: throw new NotImplementedException();
            }
            return AddText(hfText, occurrence);
        }

        public void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
        {
            if (occurrence == XLHFOccurrence.AllPages)
            {
                ClearOccurrence(XLHFOccurrence.EvenPages);
                ClearOccurrence(XLHFOccurrence.FirstPage);
                ClearOccurrence(XLHFOccurrence.OddPages);
            }
            else
            {
                ClearOccurrence(occurrence);
            }
        }

        private void ClearOccurrence(XLHFOccurrence occurrence)
        {
            if (texts.ContainsKey(occurrence))
                texts.Remove(occurrence);
        }

        private String GetHFFont(String text, IXLFont xlFont)
        {
            String retVal = String.Empty;

            retVal += xlFont.FontName != null ? "&\"" + xlFont.FontName : "\"-";
            retVal += GetHFFontBoldItalic(xlFont);
            retVal += xlFont.FontSize > 0 ? "&" + xlFont.FontSize.ToString() : "";
            retVal += xlFont.Strikethrough ? "&S" : "";
            retVal += xlFont.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript ? "&Y" : "";
            retVal += xlFont.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript ? "&X" : "";
            retVal += xlFont.Underline== XLFontUnderlineValues.Single ? "&U" : "";
            retVal += xlFont.Underline == XLFontUnderlineValues.Double ? "&E" : "";
            retVal += "&K" + xlFont.FontColor.Color.ToHex().Substring(2);

            retVal += text;

            retVal += xlFont.Underline == XLFontUnderlineValues.Double ? "&E" : "";
            retVal += xlFont.Underline == XLFontUnderlineValues.Single ? "&U" : "";
            retVal += xlFont.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript ? "&X" : "";
            retVal += xlFont.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript ? "&Y" : "";
            retVal += xlFont.Strikethrough ? "&S" : "";
            
            return retVal;
        }

        private String GetHFFontBoldItalic(IXLFont xlFont)
        {
            String retVal = String.Empty;
            if (xlFont.Bold && xlFont.Italic)
            {
                retVal += ",Bold Italic\"";
            }
            else if (xlFont.Bold)
            {
                retVal += ",Bold\"";
            }
            else if (xlFont.Italic)
            {
                retVal += ",Italic\"";
            }
            else
            {
                retVal += ",Regular\"";
            }

            return retVal;
        }
    }
}
