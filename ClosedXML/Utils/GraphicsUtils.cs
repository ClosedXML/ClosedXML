using SkiaSharp;
using System;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
        public static SKRect MeasureString(string text, SKTypeface fontName)
        {
            using var paint = new SKPaint();
            paint.Typeface = fontName;

            // Size: 12px
            paint.TextSize = 12f;

            var skBounds = SKRect.Empty;
            var textWidth = paint.MeasureText(text.AsSpan(), ref skBounds);
            return skBounds;
        }
    }
}
