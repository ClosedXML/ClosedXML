using ClosedXML.Excel.Drawings;
using System;
using System.Drawing;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// Metadata info about a picture. At least one of the sizes (pixels/physical) must be set.
    /// </summary>
    public readonly struct XLPictureInfo
    {
        /// <summary>
        /// Detected format of the image.
        /// </summary>
        public XLPictureFormat Format { get; }

        /// <summary>
        /// Size of the picture in pixels.
        /// </summary>
        public Size SizePx { get; }

        /// <summary>
        /// A physical size in 0.01mm. Used by vector images.
        /// </summary>
        public Size SizePhys { get; }

        /// <summary>
        /// 0 = use workbook DPI.
        /// </summary>
        public double DpiX { get; }

        /// <summary>
        /// 0 = use workbook DPI.
        /// </summary>
        public double DpiY { get; }

        public XLPictureInfo(XLPictureFormat format, uint width, uint height, double dpiX, double dpiY)
        {
            if (width > int.MaxValue || height > int.MaxValue)
                throw new ArgumentException("Size of picture too large.");
            Format = format;
            SizePx = new Size((int)width, (int)height);
            DpiX = dpiX;
            DpiY = dpiY;
        }

        public XLPictureInfo(XLPictureFormat format, Size sizePx, Size sizePhys) : this(format, sizePx, sizePhys, 0, 0)
        {
        }

        public XLPictureInfo(XLPictureFormat format, Size sizePx, Size sizePhys, double dpiX, double dpiY)
        {
            if (sizePx.IsEmpty && sizePhys.IsEmpty)
                throw new ArgumentException("Both sizes can't be empty.");
            Format = format;
            SizePx = sizePx;
            SizePhys = sizePhys;
            DpiX = dpiX;
            DpiY = dpiY;
        }

        internal Size GetSizePx(double dpiX, double dpiY)
        {
            if (SizePx.IsEmpty && SizePhys.IsEmpty)
                throw new InvalidOperationException("Image doesn't have a size.");

            if (!SizePx.IsEmpty)
                return SizePx;

            return new Size((int)Math.Ceiling(SizePhys.Width / 1000d / 2.54d * dpiX), (int)Math.Ceiling(SizePhys.Height / 1000d / 2.54d * dpiY));
        }
    }
}
