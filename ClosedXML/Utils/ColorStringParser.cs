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
                if (htmlColor[0] != '#')
                    htmlColor = '#' + htmlColor;

                return ColorTranslator.FromHtml(htmlColor);
            }
            catch
            {
                // https://github.com/ClosedXML/ClosedXML/issues/675
                // When regional settings list separator is # , the standard ColorTranslator.FromHtml fails
                return Color.FromArgb(int.Parse(htmlColor.Replace("#", ""), NumberStyles.AllowHexSpecifier));
            }
        }
    }
}
