using MetadataExtractor;
using MetadataExtractor.Formats.Bmp;
using MetadataExtractor.Formats.Exif;
using MetadataExtractor.Formats.Gif;
using MetadataExtractor.Formats.Ico;
using MetadataExtractor.Formats.Jpeg;
using MetadataExtractor.Formats.Pcx;
using MetadataExtractor.Formats.Png;
using MetadataExtractor.Formats.Tiff;
using MetadataExtractor.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ClosedXML.Excel.Drawings
{
    [DebuggerDisplay("{Name}")]
    internal class XLPicture : IXLPicture
    {
        private const String InvalidNameChars = @":\/?*[]";

        private static IDictionary<FileType, XLPictureFormat> FormatMap = new Dictionary<FileType, XLPictureFormat>()
        {
            [FileType.Bmp] = XLPictureFormat.Bmp,
            [FileType.Gif] = XLPictureFormat.Gif,
            [FileType.Png] = XLPictureFormat.Png,
            [FileType.Tiff] = XLPictureFormat.Tiff,
            [FileType.Ico] = XLPictureFormat.Icon,
            [FileType.Pcx] = XLPictureFormat.Pcx,
            [FileType.Jpeg] = XLPictureFormat.Jpeg
        };

        private readonly IXLWorksheet _worksheet;
        private Int32 height;
        private Int32 id;
        private String name = string.Empty;
        private Int32 width;

        internal XLPicture(IXLWorksheet worksheet, Stream stream)
            : this(worksheet)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            this.ImageStream = new MemoryStream();
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);

                DeduceImageFormat(ImageStream);
                DeduceDimensions(ImageStream);

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
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(ImageStream);
                DeduceImageFormat(ImageStream, format);
                DeduceDimensions(ImageStream);
            }
        }

#if _NETFRAMEWORK_
        internal XLPicture(IXLWorksheet worksheet, Bitmap bitmap)
            : this(worksheet)
        {
            if (bitmap == null) throw new ArgumentNullException(nameof(bitmap));
            this.ImageStream = new MemoryStream();
            bitmap.Save(ImageStream, bitmap.RawFormat);
            ImageStream.Seek(0, SeekOrigin.Begin);
            DeduceImageFormat(ImageStream);
            DeduceDimensions(ImageStream);
        }
#endif

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

            // Calculate default picture ID
            var allPictures = worksheet.Workbook.Worksheets.SelectMany(ws => ws.Pictures);
            if (allPictures.Any())
                this.id = allPictures.Max(p => p.Id) + 1;
            else
                this.id = 1;
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

        public Int32 Id
        {
            get { return id; }
            internal set
            {
                if ((_worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
                    throw new ArgumentException($"The picture ID '{value}' already exists.");
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

                if ((_worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? this) != this)
                    throw new ArgumentException($"The picture name '{value}' already exists.");

                SetName(value);
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

        internal void SetName(string value)
        {
            if (value.IndexOfAny(InvalidNameChars.ToCharArray()) != -1)
                throw new ArgumentException($"Picture names cannot contain any of the following characters: {InvalidNameChars}");

            if (String.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Picture names cannot be empty");

            if (value.Length > 31)
                throw new ArgumentException("Picture names cannot be more than 31 characters");

            name = value;
        }

        private void DeduceDimensions(Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            MetadataExtractor.Directory d;
            switch (this.Format)
            {
                case XLPictureFormat.Bmp:
                    d = BmpMetadataReader.ReadMetadata(stream);
                    this.width = d.GetInt32(BmpHeaderDirectory.TagImageWidth);
                    this.height = d.GetInt32(BmpHeaderDirectory.TagImageHeight);
                    break;

                case XLPictureFormat.Gif:
                    d = GifMetadataReader.ReadMetadata(stream).OfType<GifHeaderDirectory>().First();
                    this.width = d.GetInt32(GifHeaderDirectory.TagImageWidth);
                    this.height = d.GetInt32(GifHeaderDirectory.TagImageHeight);
                    break;

                case XLPictureFormat.Png:
                    d = PngMetadataReader.ReadMetadata(stream).OfType<PngDirectory>().First();
                    this.width = d.GetInt32(PngDirectory.TagImageWidth);
                    this.height = d.GetInt32(PngDirectory.TagImageHeight);
                    break;

                case XLPictureFormat.Tiff:
                    d = TiffMetadataReader.ReadMetadata(stream).First();
                    this.width = d.GetInt32(ExifDirectoryBase.TagImageWidth);
                    this.height = d.GetInt32(ExifDirectoryBase.TagImageHeight);
                    break;

                case XLPictureFormat.Icon:
                    d = IcoMetadataReader.ReadMetadata(stream).OfType<IcoDirectory>().First();
                    this.width = d.GetInt32(IcoDirectory.TagImageWidth);
                    this.height = d.GetInt32(IcoDirectory.TagImageHeight);
                    break;

                case XLPictureFormat.Pcx:
                    d = PcxMetadataReader.ReadMetadata(stream);
                    this.width = d.GetInt32(PcxDirectory.TagXMax);
                    this.height = d.GetInt32(PcxDirectory.TagYMax);
                    break;

                case XLPictureFormat.Jpeg:
                    d = JpegMetadataReader.ReadMetadata(stream).OfType<JpegDirectory>().First();
                    this.width = d.GetInt32(JpegDirectory.TagImageWidth);
                    this.height = d.GetInt32(JpegDirectory.TagImageHeight);
                    break;
            }
            this.OriginalWidth = this.width;
            this.OriginalHeight = this.height;
        }

        private void DeduceImageFormat(Stream stream, XLPictureFormat format)
        {
            DeduceImageFormat(stream);
            if (this.Format != format)
                throw new ArgumentException(nameof(format));
        }

        private void DeduceImageFormat(Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            var fileType = FileTypeDetector.DetectFileType(stream);
            if (fileType == FileType.Unknown)
                throw new NotImplementedException();
            else if (FormatMap.ContainsKey(fileType))
                this.Format = FormatMap[fileType];
            else
                throw new NotImplementedException();
        }
    }
}
