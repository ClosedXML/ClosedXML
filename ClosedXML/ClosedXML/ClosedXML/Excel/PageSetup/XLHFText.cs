using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLHFText
    {
        public XLHFText(IXLRichText richText)
        {
            RichText = richText;
        }
        public IXLRichText RichText { get; private set; }

        public String HFText
        {
            get
            {
                String retVal = String.Empty;

                retVal += RichText.FontName != null ? "&\"" + RichText.FontName : "\"-";
                retVal += GetHFFontBoldItalic(RichText);
                retVal += RichText.FontSize > 0 ? "&" + RichText.FontSize.ToString() : "";
                retVal += RichText.Strikethrough ? "&S" : "";
                retVal += RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript ? "&Y" : "";
                retVal += RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript ? "&X" : "";
                retVal += RichText.Underline == XLFontUnderlineValues.Single ? "&U" : "";
                retVal += RichText.Underline == XLFontUnderlineValues.Double ? "&E" : "";
                retVal += "&K" + RichText.FontColor.Color.ToHex().Substring(2);

                retVal += RichText.Text;

                retVal += RichText.Underline == XLFontUnderlineValues.Double ? "&E" : "";
                retVal += RichText.Underline == XLFontUnderlineValues.Single ? "&U" : "";
                retVal += RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript ? "&X" : "";
                retVal += RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript ? "&Y" : "";
                retVal += RichText.Strikethrough ? "&S" : "";

                return retVal;
            }
        }

        private String GetHFFontBoldItalic(IXLRichText xlFont)
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
