using System;
using Color = System.Drawing.Color;

namespace ClosedXML.Utils
{
    internal static class ColorStringParser
    {
        public static Color ParseFromArgb(string argbColor)
        {
            ReadOnlySpan<char> argb = argbColor.AsSpan();
            if (argb[0] == '#')
                argb = argb.Slice(1);

            if (argb.Length == 8)
            {
                return Color.FromArgb(
                    ReadHex(argb, 0, 2),
                    ReadHex(argb, 2, 2),
                    ReadHex(argb, 4, 2),
                    ReadHex(argb, 6, 2));
            }

            if (argb.Length == 6)
            {
                return Color.FromArgb(
                    ReadHex(argb, 0, 2),
                    ReadHex(argb, 2, 2),
                    ReadHex(argb, 4, 2));
            }

            if (argb.Length == 3)
            {
                var r = ReadHex(argb, 0, 1);
                var g = ReadHex(argb, 1, 1);
                var b = ReadHex(argb, 2, 1);
                return Color.FromArgb(
                    (r << 4) | r,
                    (g << 4) | g,
                    (b << 4) | b);
            }

            throw new FormatException($"Unable to parse color {argbColor}.");
        }

        /// <summary>
        /// Parse RRGGBB color.
        /// </summary>
        internal static Color ParseFromRgb(string rgbColor)
        {
            if (rgbColor.Length != 6)
                throw new FormatException();

            ReadOnlySpan<char> rgbSpan = rgbColor.AsSpan();

            return Color.FromArgb(
                ReadHex(rgbSpan, 0, 2),
                ReadHex(rgbSpan, 2, 2),
                ReadHex(rgbSpan, 4, 2));
        }

        private static int ReadHex(ReadOnlySpan<char> text, int start, int length)
        {
            var value = 0;
            for (var i = start; i < start + length; ++i)
            {
                var c = text[i];
                int b = c switch
                {
                    >= '0' and <= '9' => '0',
                    >= 'A' and <= 'F' => 'A' - 10,
                    >= 'a' and <= 'f' => 'a' - 10,
                    _ => throw new FormatException($"Unable to parse {text.ToString()}.")
                };
                var charValue = c - b;
                value = value * 16 + charValue;
            }

            return value;
        }
    }
}
