// Keep this file CodeMaid organised and cleaned
using DocumentFormat.OpenXml.Spreadsheet;
using SkiaSharp;

namespace ClosedXML.Excel
{
    internal static class ColorExtensions
    {
        private static readonly char[] hexDigits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F' };

        public static string ToHex(this SKColor color)
        {
            var bytes = new byte[4];

            bytes[0] = color.Alpha;

            bytes[1] = color.Red;

            bytes[2] = color.Green;

            bytes[3] = color.Blue;

            var chars = new char[bytes.Length * 2];

            for (var i = 0; i < bytes.Length; i++)
            {
                int b = bytes[i];

                chars[i * 2] = hexDigits[b >> 4];

                chars[i * 2 + 1] = hexDigits[b & 0xF];
            }

            return new string(chars);
        }

        internal static Color ConvertToOpenXmlColor(this SKColor sKColor)
        {
            var color = new Color
            {
                Rgb = new DocumentFormat.OpenXml.HexBinaryValue(sKColor.ToHex())
            };
            return color;
        }
    }
}
