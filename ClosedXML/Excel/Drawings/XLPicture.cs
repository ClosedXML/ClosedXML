using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ClosedXML.Excel.Drawings
{
    [DebuggerDisplay("{Name}")]
    internal class XLPicture : IXLPicture
    {
        private const string InvalidNameChars = @":\/?*[]";
        private const int ImageQuality = 70;
        private static readonly IDictionary<XLPictureFormat, SKEncodedImageFormat> FormatMap;

        private bool _disposed;

        private int height;
        private int id;
        private string name = string.Empty;
        private int width;
        private readonly SKCodec codec;

        static XLPicture()
        {
            var SKEncodedImageFormatValues = Enum.GetValues(typeof(SKEncodedImageFormat)).OfType<SKEncodedImageFormat>().ToList();

            FormatMap = new Dictionary<XLPictureFormat, SKEncodedImageFormat>();

            foreach (var xlpPiscutreFormat in Enum.GetValues(typeof(XLPictureFormat)))
            {
                var xlpPiscutreFormatName = xlpPiscutreFormat.ToString();
                var mathcingSKEncodedImageFormat = SKEncodedImageFormatValues.FirstOrDefault(value => value.ToString().Equals(xlpPiscutreFormatName, StringComparison.OrdinalIgnoreCase));
                FormatMap.Add((XLPictureFormat)xlpPiscutreFormat, mathcingSKEncodedImageFormat);
            }
        }

        internal XLPicture(IXLWorksheet worksheet, Stream stream) : this(worksheet)
        {
            if (stream == null) { throw new ArgumentNullException(nameof(stream)); }

            ImageStream = new MemoryStream();
            {
                stream.Position = 0;
                stream.CopyTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);

                codec = SKCodec.Create(ImageStream, out var result);

                if (codec != null)
                {
                    using var bitmap = SKBitmap.Decode(codec);
                    if (FormatMap.Values.Contains(codec.EncodedFormat))
                    {
                        Format = FormatMap.Single(f => f.Value.Equals(codec.EncodedFormat)).Key;
                    }

                    DeduceDimensionsFromBitmap(bitmap);
                }

                stream.Position = 0;
                ImageStream.Seek(0, SeekOrigin.Begin);
            }
        }

        internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
            : this(worksheet)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            Format = format;

            ImageStream = new MemoryStream();
            {
                stream.Position = 0;
                stream.CopyTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);

                codec = SKCodec.Create(ImageStream, out var result);

                if (codec != null)
                {
                    using var bitmap = SKBitmap.Decode(codec);
                    if (FormatMap.TryGetValue(Format, out var imageFormat) && imageFormat != codec.EncodedFormat)
                    {
                        throw new ArgumentException("The picture format in the stream and the parameter don't match");
                    }

                    DeduceDimensionsFromBitmap(bitmap);
                }

                ImageStream.Seek(0, SeekOrigin.Begin);
            }
        }

        internal XLPicture(IXLWorksheet worksheet, SKCodec codec) : this(worksheet)
        {
            if (codec == null)
            {
                throw new ArgumentNullException(nameof(codec));
            }

            ImageStream = new MemoryStream();

            using (var bitmap = SKBitmap.Decode(codec))
            {
                using var data = bitmap.Encode(codec.EncodedFormat, ImageQuality);
                data.SaveTo(ImageStream);
                ImageStream.Seek(0, SeekOrigin.Begin);
                DeduceDimensionsFromBitmap(bitmap);
            }

            var formats = FormatMap.Where(f => f.Value.Equals(codec.EncodedFormat));
            if (!formats.Any() || formats.Count() > 1)
            {
                throw new ArgumentException($"Unsupported or unknown image format '{codec.EncodedFormat}' in bitmap");
            }

            Format = formats.Single().Key;
        }

        private XLPicture(IXLWorksheet worksheet)
        {
            Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            Placement = XLPicturePlacement.MoveAndSize;
            Markers = new Dictionary<XLMarkerPosition, XLMarker>()
            {
                [XLMarkerPosition.TopLeft] = null,
                [XLMarkerPosition.BottomRight] = null
            };

            // Calculate default picture ID
            var allPictures = worksheet.Workbook.Worksheets.SelectMany(ws => ws.Pictures);
            if (allPictures.Any())
            {
                id = allPictures.Max(p => p.Id) + 1;
            }
            else
            {
                id = 1;
            }
        }

        public IXLCell BottomRightCell
        {
            get
            {
                ThrowIfDisposed();

                return Markers[XLMarkerPosition.BottomRight].Cell;
            }

            private set
            {
                ThrowIfDisposed();

                if (!value.Worksheet.Equals(Worksheet))
                {
                    throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");
                }

                Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
            }
        }

        public XLPictureFormat Format { get; private set; }

        public int Height
        {
            get
            {
                ThrowIfDisposed();

                return height;
            }
            set
            {
                ThrowIfDisposed();

                if (Placement == XLPicturePlacement.MoveAndSize)
                {
                    throw new ArgumentException("To set the height, the placement should be FreeFloating or Move");
                }

                height = value;
            }
        }

        public int Id
        {
            get
            {
                ThrowIfDisposed();

                return id;
            }
            internal set
            {
                ThrowIfDisposed();

                if ((Worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
                {
                    throw new ArgumentException($"The picture ID '{value}' already exists.");
                }

                id = value;
            }
        }

        public MemoryStream ImageStream { get; private set; }

        public float Left
        {
            get
            {
                ThrowIfDisposed();

                return Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0;
            }
            set
            {
                ThrowIfDisposed();

                if (Placement != XLPicturePlacement.FreeFloating)
                {
                    throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");
                }

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new SKPoint(value, Top));
            }
        }

        public string Name
        {
            get
            {
                ThrowIfDisposed();

                return name;
            }
            set
            {
                ThrowIfDisposed();

                if (name == value)
                {
                    return;
                }

                if ((Worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? this) != this)
                {
                    throw new ArgumentException($"The picture name '{value}' already exists.");
                }

                SetName(value);
            }
        }

        public int OriginalHeight { get; private set; }

        public int OriginalWidth { get; private set; }

        public XLPicturePlacement Placement { get; set; }

        public float Top
        {
            get
            {
                ThrowIfDisposed();

                return Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0;
            }
            set
            {
                ThrowIfDisposed();

                if (Placement != XLPicturePlacement.FreeFloating)
                {
                    throw new ArgumentException("To set the top offset, the placement should be FreeFloating");
                }

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new SKPoint(Left, value));
            }
        }

        public IXLCell TopLeftCell
        {
            get
            {
                ThrowIfDisposed();

                return Markers[XLMarkerPosition.TopLeft].Cell;
            }

            private set
            {
                ThrowIfDisposed();

                if (!value.Worksheet.Equals(Worksheet))
                {
                    throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");
                }

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
            }
        }

        public int Width
        {
            get
            {
                ThrowIfDisposed();

                return width;
            }
            set
            {
                ThrowIfDisposed();

                if (Placement == XLPicturePlacement.MoveAndSize)
                {
                    throw new ArgumentException("To set the width, the placement should be FreeFloating or Move");
                }

                width = value;
            }
        }

        public IXLWorksheet Worksheet { get; }

        internal IDictionary<XLMarkerPosition, XLMarker> Markers { get; private set; }

        internal string RelId { get; set; }

        /// <summary>
        /// Create a copy of the picture on a different worksheet.
        /// </summary>
        /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
        /// <returns>A created copy of the picture.</returns>
        public IXLPicture CopyTo(IXLWorksheet targetSheet)
        {
            ThrowIfDisposed();

            return CopyTo((XLWorksheet)targetSheet);
        }

        public void Delete()
        {
            ThrowIfDisposed();

            Worksheet.Pictures.Delete(Name);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                ImageStream?.Dispose();
                codec?.Dispose();
            }

            _disposed = true;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException("TemplateClass");
            }
        }

        /// <summary>
        /// Create a copy of the picture on the same worksheet.
        /// </summary>
        /// <returns>A created copy of the picture.</returns>
        public IXLPicture Duplicate()
        {
            ThrowIfDisposed();

            return CopyTo(Worksheet);
        }

        public SKPoint GetOffset(XLMarkerPosition position)
        {
            ThrowIfDisposed();

            return Markers[position].Offset;
        }

        public IXLPicture MoveTo(float left, float top)
        {
            ThrowIfDisposed();

            Placement = XLPicturePlacement.FreeFloating;
            Left = left;
            Top = top;
            return this;
        }

        public IXLPicture MoveTo(IXLCell cell)
        {
            ThrowIfDisposed();

            return MoveTo(cell, 0, 0);
        }

        public IXLPicture MoveTo(IXLCell cell, int xOffset, int yOffset)
        {
            ThrowIfDisposed();

            return MoveTo(cell, new SKPoint(xOffset, yOffset));
        }

        public IXLPicture MoveTo(IXLCell cell, SKPoint offset)
        {
            ThrowIfDisposed();

            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            Placement = XLPicturePlacement.Move;
            TopLeftCell = cell;
            Markers[XLMarkerPosition.TopLeft].Offset = offset;
            return this;
        }

        public IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell)
        {
            ThrowIfDisposed();

            return MoveTo(fromCell, 0, 0, toCell, 0, 0);
        }

        public IXLPicture MoveTo(IXLCell fromCell, int fromCellXOffset, int fromCellYOffset, IXLCell toCell, int toCellXOffset, int toCellYOffset)
        {
            ThrowIfDisposed();

            return MoveTo(fromCell, new SKPoint(fromCellXOffset, fromCellYOffset), toCell, new SKPoint(toCellXOffset, toCellYOffset));
        }

        public IXLPicture MoveTo(IXLCell fromCell, SKPoint fromOffset, IXLCell toCell, SKPoint toOffset)
        {
            ThrowIfDisposed();

            if (fromCell == null)
            {
                throw new ArgumentNullException(nameof(fromCell));
            }

            if (toCell == null)
            {
                throw new ArgumentNullException(nameof(toCell));
            }

            Placement = XLPicturePlacement.MoveAndSize;

            TopLeftCell = fromCell;
            Markers[XLMarkerPosition.TopLeft].Offset = fromOffset;

            BottomRightCell = toCell;
            Markers[XLMarkerPosition.BottomRight].Offset = toOffset;

            return this;
        }

        public IXLPicture Scale(double factor, bool relativeToOriginal = false)
        {
            ThrowIfDisposed();

            return ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);
        }

        public IXLPicture ScaleHeight(double factor, bool relativeToOriginal = false)
        {
            ThrowIfDisposed();

            Height = Convert.ToInt32((relativeToOriginal ? OriginalHeight : Height) * factor);
            return this;
        }

        public IXLPicture ScaleWidth(double factor, bool relativeToOriginal = false)
        {
            ThrowIfDisposed();

            Width = Convert.ToInt32((relativeToOriginal ? OriginalWidth : Width) * factor);
            return this;
        }

        public IXLPicture WithPlacement(XLPicturePlacement value)
        {
            ThrowIfDisposed();

            Placement = value;
            return this;
        }

        public IXLPicture WithSize(int width, int height)
        {
            ThrowIfDisposed();

            Width = width;
            Height = height;
            return this;
        }

        internal IXLPicture CopyTo(XLWorksheet targetSheet)
        {
            ThrowIfDisposed();

            if (targetSheet == null)
            {
                targetSheet = Worksheet as XLWorksheet;
            }

            IXLPicture newPicture;
            if (targetSheet == Worksheet)
            {
                newPicture = targetSheet.AddPicture(ImageStream, Format);
            }
            else
            {
                newPicture = targetSheet.AddPicture(ImageStream, Format, Name);
            }

            newPicture = newPicture
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .WithSize(Width, Height)
                    .WithPlacement(Placement);

            switch (Placement)
            {
                case XLPicturePlacement.FreeFloating:
                    newPicture.MoveTo(Left, Top);
                    break;

                case XLPicturePlacement.Move:
                    newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), GetOffset(XLMarkerPosition.TopLeft));
                    break;

                case XLPicturePlacement.MoveAndSize:
                    newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), GetOffset(XLMarkerPosition.TopLeft), targetSheet.Cell(BottomRightCell.Address),
                        GetOffset(XLMarkerPosition.BottomRight));
                    break;
            }

            return newPicture;
        }

        internal void SetName(string value)
        {
            ThrowIfDisposed();

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("Picture names cannot be empty");
            }

            if (value.IndexOfAny(InvalidNameChars.ToCharArray()) != -1)
            {
                throw new ArgumentException($"Picture names cannot contain any of the following characters: {InvalidNameChars}");
            }

            if (value.Length > 31)
            {
                throw new ArgumentException("Picture names cannot be more than 31 characters");
            }

            name = value;
        }

        private void DeduceDimensionsFromBitmap(SKBitmap bitmap)
        {
            OriginalWidth = bitmap.Width;
            OriginalHeight = bitmap.Height;

            width = bitmap.Width;
            height = bitmap.Height;
        }
    }
}