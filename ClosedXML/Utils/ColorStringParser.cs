using System;
using Color = System.Drawing.Color;

namespace ClosedXML.Utils
{
    internal static class ColorStringParser
    {
        internal static Color ParseFromHtml(string argbColor)
        {
            // Half working incorrect parser:
            // * accepts #aarrggbb, but HTML would expect #rrggbbaa
            // * doesn't accept color names
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
        /// Parse ARGB color stored in <c>ST_UnsignedIntHex</c> the same way as Excel does.
        /// </summary>
        internal static Color ParseFromArgb(ReadOnlySpan<char> argb)
        {
            // This algorithm mimics how Excel parses <c>ST_UnsignedIntHex</c> to color.
            // ST_UnsignedIntHex should be exactly 8 digits and Excel uses black for longer texts.
            if (argb.Length > 8)
                return Color.Black;

            // Excel tries to parse hex numbers as long as possible and shifts them,
            // e.g. 'ABC+' is turned into 'FF000ABC'. Signed shift keeps highest bit,
            // so keep color in uint.
            uint color = 0x00000000;
            var index = 0;
            while (index < argb.Length && TryGetHex(argb[index], out var hexDigit))
            {
                color = (color << 4) | hexDigit;
                index++;
            }

            // Although Excel always uses FF for alpha, keep alpha for valid AARRGGBB.
            var isValidArgb = index == 8;
            if (!isValidArgb)
                color |= 0xFF000000;

            return Color.FromArgb(unchecked((int)color));
        }

        /// <summary>
        /// Parse RRGGBB color.
        /// </summary>
        internal static Color ParseFromRgb(string rgbColor)
        {
            if (rgbColor.Length != 6)
                throw new FormatException("Color should have 6 chars.");

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
                if (!TryGetHex(text[i], out var hexDigit))
                    throw new FormatException($"Unable to parse {text.ToString()}.");

                value = value * 16 + (int)hexDigit;
            }

            return value;
        }

        private static bool TryGetHex(char c, out uint hexDigit)
        {
            switch (c)
            {
                case >= '0' and <= '9':
                    hexDigit = c - (uint)'0';
                    return true;
                case >= 'A' and <= 'F':
                    hexDigit = c - (uint)'A' + 10;
                    return true;
                case >= 'a' and <= 'f':
                    hexDigit = c - (uint)'a' + 10;
                    return true;
                default:
                    hexDigit = 0;
                    return false;
            }
        }
    }
}
