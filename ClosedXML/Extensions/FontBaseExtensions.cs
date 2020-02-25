// Keep this file CodeMaid organised and cleaned
using ClosedXML.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;

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

        public static Double GetHeight(this IXLFontBase fontBase, Dictionary<IXLFontBase, Font> fontCache)
        {
            var font = GetCachedFont(fontBase, fontCache);
            var textHeight = GraphicsUtils.MeasureString("X", font).Height;
            return (double)textHeight * 0.85;
        }

        public static Double GetWidth(this IXLFontBase fontBase, String text, Dictionary<IXLFontBase, Font> fontCache)
        {
            if (String.IsNullOrWhiteSpace(text))
                return 0;

            var font = GetCachedFont(fontBase, fontCache);
            var textWidth = GraphicsUtils.MeasureString(text, font).Width;

            double width = (textWidth / 7d * 256 - 128 / 7) / 256;
            width = Math.Round(width + 0.2, 2);

            return width;
        }

        private static Font GetCachedFont(IXLFontBase fontBase, Dictionary<IXLFontBase, Font> fontCache)
        {
            if (!fontCache.TryGetValue(fontBase, out Font font))
            {
                font = new Font(fontBase.FontName, (float)fontBase.FontSize, GetFontStyle(fontBase));
                fontCache.Add(fontBase, font);
            }
            return font;
        }

        private static FontStyle GetFontStyle(IXLFontBase font)
        {
            FontStyle fontStyle = FontStyle.Regular;
            if (font.Bold) fontStyle |= FontStyle.Bold;
            if (font.Italic) fontStyle |= FontStyle.Italic;
            if (font.Strikethrough) fontStyle |= FontStyle.Strikeout;
            if (font.Underline != XLFontUnderlineValues.None) fontStyle |= FontStyle.Underline;
            return fontStyle;
        }
    }
}
