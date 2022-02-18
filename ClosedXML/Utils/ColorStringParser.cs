using SkiaSharp;
using System;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Utils
{
    internal static class ColorStringParser
    {
        public static SKColor ParseFromHtml(string htmlColor)
        {
            return FromArgb(int.Parse(htmlColor.Replace("#", ""), NumberStyles.AllowHexSpecifier));
        }

        internal static SKColor FromArgb(int r, int g, int b)
        {
            return new SKColor(Convert.ToByte(r), Convert.ToByte(g), Convert.ToByte(b));
        }

        public static SKColor FromArgb(int argb)
        {
            return new SKColor((uint)argb);
        }

        internal static SKColor FromArgb(int a, int r, int g, int b)
        {
            return new SKColor(Convert.ToByte(r), Convert.ToByte(g), Convert.ToByte(b), Convert.ToByte(a));
        }

        internal static SKColor FromName(string name)
        {
            var value = (SKColor)typeof(SKColors)
                 .GetFields(BindingFlags.Static | BindingFlags.Public).Single(color => color.Name == name).GetValue(null);

            return value;
        }
    }
}
