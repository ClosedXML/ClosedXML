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
        private const int maxExcelColumnHeight = 409;

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
            var height = Math.Round(textHeight * 1.9921875, 2);
            return height < maxExcelColumnHeight ? height : maxExcelColumnHeight;
        }

        public static double GetWidth(this IXLFontBase fontBase, string text, Dictionary<IXLFontBase, SKFont> fontCache)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            var font = GetCachedFont(fontBase, fontCache);
            var marginPoints = ((font.Size * 0.4) + 8) / 1.326;
            var textWidthPoints = GraphicsUtils.MeasureString(text, font).Width;
            var columnWidth = (textWidthPoints + marginPoints) * 0.186;
            var width = Math.Round(columnWidth, 2);

            return width < 255 ? width : 255;
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