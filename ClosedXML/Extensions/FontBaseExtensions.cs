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
        private const int maxExcelColumnWidth = 255;
        private static double? CachedWidthCalibrationFactor;
        private static double? CachedHeightCalibrationFactor;

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
            var systemSpecificWidthScalingFactor = GetCachedHeightCalibration(fontCache);
            var height = SystemSpecificHeightCalculator(fontBase, fontCache, systemSpecificWidthScalingFactor);

            return height < maxExcelColumnHeight ? height : maxExcelColumnHeight;
        }

        private static double GetCachedHeightCalibration(Dictionary<IXLFontBase, SKFont> fontCache)
        {
            if (CachedHeightCalibrationFactor.HasValue)
                return CachedHeightCalibrationFactor.Value;

            var calibratedValue = 288;

            var xLFont = new XLFont
            {
                FontSize = 200,
                FontName = "Verdana"
            };

            var SystemSpecificWidthOfKnownWidth = SystemSpecificHeightCalculator(xLFont, fontCache, 1);
            CachedHeightCalibrationFactor = calibratedValue / SystemSpecificWidthOfKnownWidth;
            return CachedHeightCalibrationFactor.Value;
        }

        private static double SystemSpecificHeightCalculator(IXLFontBase fontBase, Dictionary<IXLFontBase, SKFont> fontCache, double systemSpecificHeightScalingFactor)
        {
            var font = GetCachedFont(fontBase, fontCache);
            var textHeight = GraphicsUtils.MeasureString("X", font).Height;
            var height = Math.Round(textHeight * systemSpecificHeightScalingFactor, 2);
            return height;
        }

        public static double GetWidth(this IXLFontBase fontBase, string text, Dictionary<IXLFontBase, SKFont> fontCache)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }
            var systemSpecificWidthScalingFactor = GetCachedWidthCalibration(fontCache);
            var width = SystemSpecificWidthCalculator(fontBase, text, fontCache, systemSpecificWidthScalingFactor);

            return width < maxExcelColumnWidth ? width : maxExcelColumnWidth;
        }

        private static double GetCachedWidthCalibration(Dictionary<IXLFontBase, SKFont> fontCache)
        {
            if (CachedWidthCalibrationFactor.HasValue)
                return CachedWidthCalibrationFactor.Value;

            var calibratedValue = 36.535187641402715d;
            var text = "Very Wide Column";

            var xLFont = new XLFont
            {
                FontSize = 20,
                FontName = "Verdana"
            };

            var SystemSpecificWidthOfKnownWidth = SystemSpecificWidthCalculator(xLFont, text, fontCache, 1);
            CachedWidthCalibrationFactor = calibratedValue / SystemSpecificWidthOfKnownWidth;

            return CachedWidthCalibrationFactor.Value;
        }

        private static double SystemSpecificWidthCalculator(IXLFontBase fontBase, string text, Dictionary<IXLFontBase, SKFont> fontCache, double systemSpecificScalingFactor)
        {
            var font = GetCachedFont(fontBase, fontCache);
            var marginPoints = ((font.Size * 0.4) + 8) / 1.326;
            var textWidthPoints = GraphicsUtils.MeasureString(text, font).Width;
            var columnWidth = (textWidthPoints + marginPoints) * systemSpecificScalingFactor;
            var width = Math.Round(columnWidth, 2);
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