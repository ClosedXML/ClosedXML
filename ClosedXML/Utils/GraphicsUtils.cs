using System;
using System.Drawing;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
#if _NETFRAMEWORK_
        [ThreadStatic]
        private static Graphics threadLocalGraphics;
        internal static Graphics Graphics
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
            SizeF result = Graphics.MeasureString(s, font, Int32.MaxValue, StringFormat.GenericTypographic);
            return result;
        }
#else

        internal static Graphics Graphics = new Graphics();
#endif
    }

#if _NETSTANDARD_

    // Stub structure for .NET Standard
    internal struct Graphics
    {
        public float DpiX { get { return 96; } }
        public float DpiY { get { return 96; } }
    }

#endif
}
