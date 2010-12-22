using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLHFItem : IXLHFItem
    {
        private Dictionary<XLHFOccurrence, String> texts = new Dictionary<XLHFOccurrence, String>();
        public String GetText(XLHFOccurrence occurrence)
        {
            if(texts.ContainsKey(occurrence))
                return texts[occurrence];
            else
                return String.Empty;
        }

        public void AddText(String text)
        {
            AddText(text, XLHFOccurrence.AllPages);
        }
        public void AddText(XLHFPredefinedText predefinedText)
        {
            AddText(predefinedText, XLHFOccurrence.AllPages);
        }
        public void AddText(String text, XLHFOccurrence occurrence)
        {
            AddText(text, occurrence, null);
        }
        public void AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence)
        {
            AddText(predefinedText, occurrence, null);
        }

        public void AddText(String text, XLHFOccurrence occurrence, IXLFont xlFont)
        {
            if (text.Length > 0)
            {
                var newText = xlFont != null ? GetHFFont(text, xlFont) : text;
                //var newText = hfFont + text;
                if (occurrence == XLHFOccurrence.AllPages)
                {
                    AddTextToOccurrence(newText, XLHFOccurrence.EvenPages);
                    AddTextToOccurrence(newText, XLHFOccurrence.FirstPage);
                    AddTextToOccurrence(newText, XLHFOccurrence.OddPages);
                }
                else
                {
                    AddTextToOccurrence(newText, occurrence);
                }
            }
        }

        private void AddTextToOccurrence(String text, XLHFOccurrence occurrence)
        {
            if (text.Length > 0)
            {
                var newText = text;
                if (texts.ContainsKey(occurrence))
                    texts[occurrence] = texts[occurrence] + newText;
                else
                    texts.Add(occurrence, newText);
            }
        }

        public void AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence, IXLFont xlFont)
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
            AddText(hfText, occurrence, xlFont);
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
