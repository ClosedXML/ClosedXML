// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Utils;
using SkiaSharp;
using System;
using System.Collections.Generic;

namespace ClosedXML.Extensions
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

        public static double GetHeight(this IXLFontBase fontBase, Dictionary<IXLFontBase, SKFont> fontCache)
        {
            var font = GetCachedFont(fontBase, fontCache);
            var textHeight = GraphicsUtils.MeasureString("X", font).Height;
            return (double)textHeight * 1.8;
        }

        public static double GetWidth(this IXLFontBase fontBase, string text, Dictionary<IXLFontBase, SKFont> fontCache)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            var font = GetCachedFont(fontBase, fontCache);
            var textWidth = GraphicsUtils.MeasureString(text, font).Width;

            var width = ((textWidth / 7d * 256) - (128 / 7)) / 256;
            width = Math.Round(width + 1.2, 2);

            return width;
        }

        private static SKFont GetCachedFont(IXLFontBase fontBase, Dictionary<IXLFontBase, SKFont> fontCache)
        {
            if (!fontCache.TryGetValue(fontBase, out var font))
            {
                using var fontManager = SKFontManager.CreateDefault();
                var typeface = fontManager.MatchFamily(fontBase.FontName);
                font = new SKFont(typeface, (float)fontBase.FontSize);
                fontCache.Add(fontBase, font);
            }
            return font;
        }
    }
}