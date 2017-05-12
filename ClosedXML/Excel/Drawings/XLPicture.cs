using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace ClosedXML.Excel.Drawings
{
    public class XLPicture : IXLPicture
    {
        private MemoryStream imgStream;
        private List<IXLMarker> Markers;
        private String name;
        private bool isAbsolute;
        private ImagePartType type = ImagePartType.Jpeg;

        private long iMaxWidth = 500;
        private long iMaxHeight = 500;

        private long iWidth;
        private long iHeight;

        private long iOffsetX;
        private long iOffsetY;

        private float iVerticalResolution;
        private float iHorizontalResolution;

        private bool isResized = false;

        private void Resize()
        {
            if (iWidth > iMaxHeight || iHeight > iMaxWidth)
            {
                var scaleX = (double)iWidth / (double)iMaxWidth;
                var scaleY = (double)iHeight / (double)iMaxHeight;
                var scale = Math.Max(scaleX, scaleY);
                iWidth = (int)((double)iWidth / scale);
                iHeight = (int)((double)iHeight / scale);
            }
            isResized = true;
        }

        public long MaxWidth
        {
            get
            {
                return ConvertToEmu(iMaxWidth, iHorizontalResolution);
            }
            set
            {
                iMaxWidth = value;
                isResized = false;
            }
        }


        public long MaxHeight
        {
            get
            {
                return ConvertToEmu(iMaxHeight, iVerticalResolution);
            }
            set
            {
                iMaxHeight = value;
                isResized = false;
            }
        }

        public long Width
        {
            get
            {
                if (!isResized)
                {
                    Resize();
                }
                return ConvertToEmu(iWidth, iHorizontalResolution);
            }
            set { }
        }

        public long Height
        {
            get
            {
                if (!isResized)
                {
                    Resize();
                }
                return ConvertToEmu(iHeight, iVerticalResolution);
            }
            set { }
        }

        public long RawHeight
        {
            get { return (long)iHeight; }
        }
        public long RawWidth
        {
            get { return (long)iWidth; }
        }

        public long OffsetX
        {
            get { return ConvertToEmu(iOffsetX, iHorizontalResolution); }
            set { iOffsetX = value; }
        }
        public long OffsetY
        {
            get { return ConvertToEmu(iOffsetY, iVerticalResolution); }
            set { iOffsetY = value; }
        }

        public long RawOffsetX
        {
            get
            {
                return iOffsetX;
            }
            set
            {
                iOffsetX = value;
            }
        }

        public long RawOffsetY
        {
            get
            {
                return iOffsetY;
            }
            set
            {
                iOffsetY = value;
            }
        }

        private long ConvertToEmu(long pixels, float resolution)
        {
            return (long)(914400 * pixels / resolution);
        }

        public Stream ImageStream
        {
            get
            {
                return imgStream;
            }
            set
            {
                if (imgStream == null)
                {
                    imgStream = new MemoryStream();
                }
                else
                {
                    imgStream.Dispose();
                    imgStream = new MemoryStream();
                }
                value.CopyTo(imgStream);
                imgStream.Seek(0, SeekOrigin.Begin);

                using (var bitmap = new System.Drawing.Bitmap(imgStream))
                {
                    iWidth = (long)bitmap.Width;
                    iHeight = (long)bitmap.Height;
                    iHorizontalResolution = bitmap.HorizontalResolution;
                    iVerticalResolution = bitmap.VerticalResolution;
                }
                imgStream.Seek(0, SeekOrigin.Begin);
            }
        }

        public List<IXLMarker> GetMarkers()
        {
            return Markers != null ? Markers : new List<IXLMarker>();
        }
        public void AddMarker(IXLMarker marker)
        {
            if (Markers == null)
            {
                Markers = new List<IXLMarker>();
            }
            Markers.Add(marker);
        }

        public String Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        public bool IsAbsolute
        {
            get
            {
                return isAbsolute;
            }
            set
            {
                isAbsolute = value;
            }
        }

        public String Type
        {
            get
            {
                return GetExtension(type);
            }
            set
            {
                try
                {
                    type = (ImagePartType)Enum.Parse(typeof(ImagePartType), value, true);
                }
                catch
                {
                    type = ImagePartType.Jpeg;
                }
            }
        }

        private String GetExtension(ImagePartType type)
        {
            return type.ToString().ToLower();
        }

        public ImagePartType GetImagePartType()
        {
            return type;
        }
    }
}
