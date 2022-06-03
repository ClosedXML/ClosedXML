using System;
using System.Drawing;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
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
    }
}
