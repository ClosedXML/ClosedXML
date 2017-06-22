using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Excel.Drawings
{
    [DebuggerDisplay("{Name}")]
    internal class XLPicture : IXLPicture
    {
        private const String InvalidNameChars = @":\/?*[]";
        private static IDictionary<XLPictureFormat, ImageFormat> FormatMap;
        private readonly IXLWorksheet _worksheet;
        private Int32 height;
        private String name = string.Empty;
        private Int32 width;

        static XLPicture()
        {
            var properties = typeof(ImageFormat).GetProperties(BindingFlags.Static | BindingFlags.Public);
            FormatMap = Enum.GetValues(typeof(XLPictureFormat))
                .Cast<XLPictureFormat>()
                .Where(pf => properties.Any(pi => pi.Name.Equals(pf.ToString(), StringComparison.OrdinalIgnoreCase)))
                .ToDictionary(
                    pf => pf,
                    pf => properties.Single(pi => pi.Name.Equals(pf.ToString(), StringComparison.OrdinalIgnoreCase)).GetValue(null, null) as ImageFormat
                );
        }

        internal XLPicture(IXLWorksheet worksheet, Stream stream)
                    : this(worksheet)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            this.ImageStream = new MemoryStream();
            {
                stream.Position = 0;
                stream.CopyTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);

                using (var bitmap = new Bitmap(ImageStream))
                {
                    if (FormatMap.Values.Select(f => f.Guid).Contains(bitmap.RawFormat.Guid))
                        this.Format = FormatMap.Single(f => f.Value.Guid.Equals(bitmap.RawFormat.Guid)).Key;

                    DeduceDimensionsFromBitmap(bitmap);
                }
                ImageStream.Seek(0, SeekOrigin.Begin);
            }
        }

        internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
            : this(worksheet)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            this.Format = format;

            this.ImageStream = new MemoryStream();
            {
                stream.Position = 0;
                stream.CopyTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);

                using (var bitmap = new Bitmap(ImageStream))
                {
                    if (FormatMap.ContainsKey(this.Format))
                    {
                        if (FormatMap[this.Format].Guid != bitmap.RawFormat.Guid)
                            throw new ArgumentException("The picture format in the stream and the parameter don't match");
                    }

                    DeduceDimensionsFromBitmap(bitmap);
                }
                ImageStream.Seek(0, SeekOrigin.Begin);
            }
        }

        internal XLPicture(IXLWorksheet worksheet, Bitmap bitmap)
            : this(worksheet)
        {
            if (bitmap == null) throw new ArgumentNullException(nameof(bitmap));
            this.ImageStream = new MemoryStream();
            bitmap.Save(ImageStream, bitmap.RawFormat);
            ImageStream.Seek(0, SeekOrigin.Begin);
            DeduceDimensionsFromBitmap(bitmap);

            var formats = FormatMap.Where(f => f.Value.Guid.Equals(bitmap.RawFormat.Guid));
            if (!formats.Any() || formats.Count() > 1)
                throw new ArgumentException("Unsupported or unknown image format in bitmap");

            this.Format = formats.Single().Key;
        }

        private XLPicture(IXLWorksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            this._worksheet = worksheet;
            this.Placement = XLPicturePlacement.MoveAndSize;
            this.Markers = new Dictionary<XLMarkerPosition, IXLMarker>()
            {
                [XLMarkerPosition.TopLeft] = null,
                [XLMarkerPosition.BottomRight] = null
            };
        }

        public IXLAddress BottomRightCellAddress
        {
            get
            {
                return Markers[XLMarkerPosition.BottomRight].Address;
            }

            private set
            {
                if (!value.Worksheet.Equals(this._worksheet))
                    throw new ArgumentOutOfRangeException(nameof(value.Worksheet));
                this.Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
            }
        }

        public XLPictureFormat Format { get; private set; }

        public Int32 Height
        {
            get { return height; }
            set
            {
                if (this.Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the height, the placement should be FreeFloating or Move");
                height = value;
            }
        }

        public MemoryStream ImageStream { get; private set; }

        public Int32 Left
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(_worksheet.Cell(1, 1).Address, new Point(value, this.Top));
            }
        }

        public String Name
        {
            get { return name; }
            set
            {
                if (name == value) return;

                if (value.IndexOfAny(InvalidNameChars.ToCharArray()) != -1)
                    throw new ArgumentException($"Picture names cannot contain any of the following characters: {InvalidNameChars}");

                if (XLHelper.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Picture names cannot be empty");

                if (value.Length > 31)
                    throw new ArgumentException("Picture names cannot be more than 31 characters");

                if ((_worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? this) != this)
                    throw new ArgumentException($"The picture name '{value}' already exists.");

                name = value;
            }
        }

        public Int32 OriginalHeight { get; private set; }

        public Int32 OriginalWidth { get; private set; }

        public XLPicturePlacement Placement { get; set; }

        public Int32 Top
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the top offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(_worksheet.Cell(1, 1).Address, new Point(this.Left, value));
            }
        }

        public IXLAddress TopLeftCellAddress
        {
            get
            {
                return Markers[XLMarkerPosition.TopLeft].Address;
            }

            private set
            {
                if (!value.Worksheet.Equals(this._worksheet))
                    throw new ArgumentOutOfRangeException(nameof(value.Worksheet));

                this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
            }
        }

        public Int32 Width
        {
            get { return width; }
            set
            {
                if (this.Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the width, the placement should be FreeFloating or Move");
                width = value;
            }
        }

        public IXLWorksheet Worksheet
        {
            get { return _worksheet; }
        }

        internal IDictionary<XLMarkerPosition, IXLMarker> Markers { get; private set; }

        internal String RelId { get; set; }

        public void Delete()
        {
            Worksheet.Pictures.Delete(this.Name);
        }

        public void Dispose()
        {
            this.ImageStream.Dispose();
        }

        public Point GetOffset(XLMarkerPosition position)
        {
            return Markers[position].Offset;
        }

        public IXLPicture MoveTo(Int32 left, Int32 top)
        {
            this.Placement = XLPicturePlacement.FreeFloating;
            this.Left = left;
            this.Top = top;
            return this;
        }

        public IXLPicture MoveTo(IXLAddress cell)
        {
            return MoveTo(cell, 0, 0);
        }

        public IXLPicture MoveTo(IXLAddress cell, Int32 xOffset, Int32 yOffset)
        {
            return MoveTo(cell, new Point(xOffset, yOffset));
        }

        public IXLPicture MoveTo(IXLAddress cell, Point offset)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            this.Placement = XLPicturePlacement.Move;
            this.TopLeftCellAddress = cell;
            this.Markers[XLMarkerPosition.TopLeft].Offset = offset;
            return this;
        }

        public IXLPicture MoveTo(IXLAddress fromCell, IXLAddress toCell)
        {
            return MoveTo(fromCell, 0, 0, toCell, 0, 0);
        }

        public IXLPicture MoveTo(IXLAddress fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXLAddress toCell, Int32 toCellXOffset, Int32 toCellYOffset)
        {
            return MoveTo(fromCell, new Point(fromCellXOffset, fromCellYOffset), toCell, new Point(toCellXOffset, toCellYOffset));
        }

        public IXLPicture MoveTo(IXLAddress fromCell, Point fromOffset, IXLAddress toCell, Point toOffset)
        {
            if (fromCell == null) throw new ArgumentNullException(nameof(fromCell));
            if (toCell == null) throw new ArgumentNullException(nameof(toCell));
            this.Placement = XLPicturePlacement.MoveAndSize;

            this.TopLeftCellAddress = fromCell;
            this.Markers[XLMarkerPosition.TopLeft].Offset = fromOffset;

            this.BottomRightCellAddress = toCell;
            this.Markers[XLMarkerPosition.BottomRight].Offset = toOffset;

            return this;
        }

        public IXLPicture Scale(Double factor, Boolean relativeToOriginal = false)
        {
            return this.ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);
        }

        public IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false)
        {
            this.Height = Convert.ToInt32((relativeToOriginal ? this.OriginalHeight : this.Height) * factor);
            return this;
        }

        public IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false)
        {
            this.Width = Convert.ToInt32((relativeToOriginal ? this.OriginalWidth : this.Width) * factor);
            return this;
        }

        public IXLPicture WithPlacement(XLPicturePlacement value)
        {
            this.Placement = value;
            return this;
        }

        public IXLPicture WithSize(Int32 width, Int32 height)
        {
            this.Width = width;
            this.Height = height;
            return this;
        }

        private static ImageFormat FromMimeType(string mimeType)
        {
            var guid = ImageCodecInfo.GetImageDecoders().FirstOrDefault(c => c.MimeType.Equals(mimeType, StringComparison.OrdinalIgnoreCase))?.FormatID;
            if (!guid.HasValue) return null;
            var property = typeof(System.Drawing.Imaging.ImageFormat).GetProperties(BindingFlags.Public | BindingFlags.Static)
                .FirstOrDefault(pi => (pi.GetValue(null, null) as ImageFormat).Guid.Equals(guid.Value));

            if (property == null) return null;
            return (property.GetValue(null, null) as ImageFormat);
        }

        private static string GetMimeType(Image i)
        {
            var imgguid = i.RawFormat.Guid;
            foreach (ImageCodecInfo codec in ImageCodecInfo.GetImageDecoders())
            {
                if (codec.FormatID == imgguid)
                    return codec.MimeType;
            }
            return "image/unknown";
        }

        private void DeduceDimensionsFromBitmap(Bitmap bitmap)
        {
            this.OriginalWidth = bitmap.Width;
            this.OriginalHeight = bitmap.Height;

            this.width = bitmap.Width;
            this.height = bitmap.Height;
        }
    }
}
