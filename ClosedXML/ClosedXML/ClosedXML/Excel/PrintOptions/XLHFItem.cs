using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLHFItem: IXLHFItem
    {
        private Dictionary<XLHFOccurrence, String> texts = new Dictionary<XLHFOccurrence, String>();
        public String GetText(XLHFOccurrence occurrence)
        {
            if(texts.ContainsKey(occurrence))
                return texts[occurrence];
            else
                return String.Empty;
        }

        public void AddText(String text, XLHFOccurrence occurrence = XLHFOccurrence.AllPages, IXLFont xlFont = null)
        {
            if (text.Length > 0)
            {
                var hfFont = GetHFFont(xlFont);
                var newText = hfFont + text;
                if (texts.ContainsKey(occurrence))
                    texts[occurrence] = texts[occurrence] + newText;
                else
                    texts.Add(occurrence, newText);
            }
        }

        public void AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence = XLHFOccurrence.AllPages, IXLFont xlFont = null)
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
                default: throw new NotImplementedException();
            }
            AddText(hfText, occurrence, xlFont);
        }

        public void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
        {
            if (texts.ContainsKey(occurrence))
                texts.Remove(occurrence);
        }

        private String GetHFFont(IXLFont xlFont)
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
            retVal += "&K" + xlFont.FontColor.ToHex();
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
