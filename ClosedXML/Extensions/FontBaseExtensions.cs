// Keep this file CodeMaid organised and cleaned
using ClosedXML.Utils;
using System;

namespace ClosedXML.Excel
{
    internal static class FontBaseExtensions
    {
        public static void CopyFont(this IXLFontBase font, IXLFontBase sourceFont)
        {
            font.Bold = sourceFont.Bold;
            font.Italic = sourceFont.Italic;
            font.Underline = sourceFont.Underline;
            font.Strikethrough = sourceFont.Strikethrough;
            font.VerticalAlignment = sourceFont.VerticalAlignment;
            font.Shadow = sourceFont.Shadow;
            font.FontSize = sourceFont.FontSize;
            font.FontColor = sourceFont.FontColor;
            font.FontName = sourceFont.FontName;
            font.FontFamilyNumbering = sourceFont.FontFamilyNumbering;
            font.FontCharSet = sourceFont.FontCharSet;
        }

        public static Double GetHeight(this IXLFontBase fontBase, FontCache fontCache)
        {
            var textHeight = fontCache.MeasureString("X", fontBase).Height;
            return (double)textHeight * 0.85;
        }

        public static Double GetWidth(this IXLFontBase fontBase, String text, FontCache fontCache)
        {
            if (String.IsNullOrWhiteSpace(text))
                return 0;

            var textWidth = fontCache.MeasureString(text, fontBase).Width;

            double width = (textWidth / 7d * 256 - 128 / 7) / 256;
            width = Math.Round(width + 0.2, 2);

            return width;
        }
    }
}
