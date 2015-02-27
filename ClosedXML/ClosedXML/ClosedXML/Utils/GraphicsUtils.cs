using System;
using System.Drawing;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
        [ThreadStatic]
        private static Graphics threadLocalGraphics;
        private static Graphics g
        {
            get
            {
                if (threadLocalGraphics == null)
                {
                    var image = new Bitmap(1, 1);
                    threadLocalGraphics = Graphics.FromImage(image);
                }
                return threadLocalGraphics;
            }
        }

        public static SizeF MeasureString(string s, Font font)
        {
            SizeF result = g.MeasureString(s, font, Int32.MaxValue, StringFormat.GenericTypographic);
            return result;
        }
    }
}
