using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLHFText
    {
        private readonly XLHFItem _hfItem;

        public XLHFText(XLRichString richText, XLHFItem hfItem)
        {
            RichText = richText;
            _hfItem = hfItem;
        }

        public XLRichString RichText { get; private set; }

        public String GetHFText(String prevText)
        {
            var wsFont = _hfItem.HeaderFooter.Worksheet.Style.Font;

            var isRichText = RichText.FontName != null && RichText.FontName != wsFont.FontName
                             || RichText.Bold != wsFont.Bold
                             || RichText.Italic != wsFont.Italic
                             || RichText.Strikethrough != wsFont.Strikethrough
                             || RichText.FontSize > 0 && Math.Abs(RichText.FontSize - wsFont.FontSize) > XLHelper.Epsilon
                             || RichText.VerticalAlignment != wsFont.VerticalAlignment
                             || RichText.Underline != wsFont.Underline
                             || !RichText.FontColor.Equals(wsFont.FontColor);

            if (!isRichText)
                return RichText.Text;

            StringBuilder sb = new StringBuilder();

            if (RichText.FontName != null && RichText.FontName != wsFont.FontName)
                sb.Append("&\"" + RichText.FontName);
            else
                sb.Append("&\"-");

            if (RichText.Bold && RichText.Italic)
                sb.Append(",Bold Italic\"");
            else if (RichText.Bold)
                sb.Append(",Bold\"");
            else if (RichText.Italic)
                sb.Append(",Italic\"");
            else
                sb.Append(",Regular\"");

            if (RichText.FontSize > 0 && Math.Abs(RichText.FontSize - wsFont.FontSize) > XLHelper.Epsilon)
                sb.Append("&" + RichText.FontSize);

            if (RichText.Strikethrough && !wsFont.Strikethrough)
                sb.Append("&S");

            if (RichText.VerticalAlignment != wsFont.VerticalAlignment)
            {
                if (RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript)
                    sb.Append("&Y");
                else if (RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript)
                    sb.Append("&X");
            }

            if (RichText.Underline != wsFont.Underline)
            {
                if (RichText.Underline == XLFontUnderlineValues.Single)
                    sb.Append("&U");
                else if (RichText.Underline == XLFontUnderlineValues.Double)
                    sb.Append("&E");
            }

            var lastColorPosition = prevText.LastIndexOf("&K");

            if (
                (lastColorPosition >= 0 && !RichText.FontColor.Equals(XLColor.FromHtml("#" + prevText.Substring(lastColorPosition + 2, 6))))
                || (lastColorPosition == -1 && !RichText.FontColor.Equals(wsFont.FontColor))
                )
                sb.Append("&K" + RichText.FontColor.Color.ToHex().Substring(2));

            sb.Append(RichText.Text);

            if (RichText.Underline != wsFont.Underline)
            {
                if (RichText.Underline == XLFontUnderlineValues.Single)
                    sb.Append("&U");
                else if (RichText.Underline == XLFontUnderlineValues.Double)
                    sb.Append("&E");
            }

            if (RichText.VerticalAlignment != wsFont.VerticalAlignment)
            {
                if (RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript)
                    sb.Append("&Y");
                else if (RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript)
                    sb.Append("&X");
            }

            if (RichText.Strikethrough && !wsFont.Strikethrough)
                sb.Append("&S");

            return sb.ToString();
        }
    }
}
