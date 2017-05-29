using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    internal class XLPicture : IXLPicture
    {
        internal readonly float HorizontalResolution;
        internal readonly float VerticalResolution;

        internal XLPicture(Stream stream, XLPictureFormat format)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            this.Format = format;

            this.ImageStream = new MemoryStream();
            {
                stream.CopyTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);

                using (var bitmap = new Bitmap(ImageStream))
                {
                    var expectedFormat = typeof(System.Drawing.Imaging.ImageFormat).GetProperty(this.Format.ToString()).GetValue(null, null) as System.Drawing.Imaging.ImageFormat;
                    if (expectedFormat.Guid != bitmap.RawFormat.Guid)
                        throw new ArgumentException("The picture format in the stream and the parameter don't match");

                    this.Width = bitmap.Width;
                    this.Height = bitmap.Height;
                    HorizontalResolution = bitmap.HorizontalResolution;
                    VerticalResolution = bitmap.VerticalResolution;
                }
                ImageStream.Seek(0, SeekOrigin.Begin);
            }
        }

        public XLPictureFormat Format { get; protected set; }

        public long Height { get; set; }

        public MemoryStream ImageStream { get; protected set; }

        public bool IsAbsolute { get; private set; }

        public long Left { get; set; }
        public IList<IXLMarker> Markers { get; private set; } = new List<IXLMarker>();
        public String Name { get; set; }

        public long Top { get; set; }
        public long Width { get; set; }

        public IXLPicture AtPosition(long left, long top)
        {
            this.Left = left;
            this.Top = top;
            return this;
        }

        public void Dispose()
        {
            this.ImageStream.Dispose();
        }

        public IXLPicture SetAbsolute()
        {
            return SetAbsolute(true);
        }

        public IXLPicture SetAbsolute(bool value)
        {
            this.IsAbsolute = value;
            return this;
        }

        public IXLMarker WithMarker(IXLMarker marker)
        {
            if (marker == null) throw new ArgumentNullException(nameof(marker));
            this.Markers.Add(marker);
            return marker;
        }

        public IXLPicture WithSize(long width, long height)
        {
            this.Width = width;
            this.Height = height;
            return this;
        }
    }
}
