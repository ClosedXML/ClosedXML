using System;
using System.Drawing;
using System.Globalization;

namespace ClosedXML.Utils
{
    internal static class ColorStringParser
    {
        public static Color ParseFromArgb(string argbColor)
        {
            if (argbColor[0] == '#')
                argbColor = argbColor.Substring(1);

            if (argbColor.Length == 8)
            {
                return Color.FromArgb(
                    ReadHex(argbColor, 0, 2),
                    ReadHex(argbColor, 2, 2),
                    ReadHex(argbColor, 4, 2),
                    ReadHex(argbColor, 6, 2));
            }

            if (argbColor.Length == 6)
            {
                return Color.FromArgb(
                    ReadHex(argbColor, 0, 2),
                    ReadHex(argbColor, 2, 2),
                    ReadHex(argbColor, 4, 2));
            }

            if (argbColor.Length == 3)
            {
                var r = ReadHex(argbColor, 0, 1);
                var g = ReadHex(argbColor, 1, 1);
                var b = ReadHex(argbColor, 2, 1);
                return Color.FromArgb(
                    (r << 4) | r,
                    (g << 4) | g,
                    (b << 4) | b);
            }

            throw new FormatException($"Unable to parse color {argbColor}.");
        }

        private static int ReadHex(string text, int start, int length)
        {
            return Int32.Parse(text.Substring(start, length), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        }
    }
}
