using SkiaSharp;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
        internal static SKRect MeasureString(string text, SKFont font)
        {
            using var paint = new SKPaint();
            paint.Typeface = font.Typeface;
            paint.TextSize = font.Size;

            var skBounds = SKRect.Empty;
            paint.MeasureText(text, ref skBounds);
            return skBounds;
        }
    }
}