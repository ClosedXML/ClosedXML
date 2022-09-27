using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;

namespace ClosedXML.Utils
{
    internal static class ColorStringParser
    {
        public static Color ParseFromHtml(string htmlColor)
        {
            try
            {
                if (htmlColor[0] == '#' && (htmlColor.Length == 4 || htmlColor.Length == 7))
                {
                    if (htmlColor.Length == 4)
                    {
                        var r = ReadHex(htmlColor, 1, 1);
                        var g = ReadHex(htmlColor, 2, 1);
                        var b = ReadHex(htmlColor, 3, 1);
                        return Color.FromArgb(
                            (r << 4) | r,
                            (g << 4) | g,
                            (b << 4) | b);
                    }

                    return Color.FromArgb(
                        ReadHex(htmlColor, 1, 2),
                        ReadHex(htmlColor, 3, 2),
                        ReadHex(htmlColor, 5, 2));
                }

                return (Color)TypeDescriptor.GetConverter(typeof(Color)).ConvertFromString(htmlColor);
            }
            catch
            {
                // https://github.com/ClosedXML/ClosedXML/issues/675
                // When regional settings list separator is # , the standard ColorTranslator.FromHtml fails
                return Color.FromArgb(int.Parse(htmlColor.Replace("#", ""), NumberStyles.AllowHexSpecifier));
            }
        }

        private static int ReadHex(string text, int start, int length)
        {
            return Convert.ToInt32(text.Substring(start, length), 16);
        }
    }
}
