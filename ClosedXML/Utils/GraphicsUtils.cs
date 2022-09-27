using System;
using System.Drawing;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
        [ThreadStatic]
        private static System.Drawing.Graphics threadLocalGraphics;

        internal static System.Drawing.Graphics Graphics
        {
            get
            {
                if (threadLocalGraphics == null)
                {
                    var image = new Bitmap(1, 1);
                    threadLocalGraphics = System.Drawing.Graphics.FromImage(image);
                }
                return threadLocalGraphics;
            }
        }

        private static StringFormat defaultStringFormat = StringFormat.GenericTypographic;
        public static SizeF MeasureString(string s, Font font)
        {
            SizeF result = Graphics.MeasureString(s, font, Int32.MaxValue, defaultStringFormat);
            return result;
        }
    }
}
