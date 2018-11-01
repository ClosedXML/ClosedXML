// Keep this file CodeMaid organised and cleaned
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
        private IXLMeasure height;
        private Int32 id;
        private String name = string.Empty;
        private IXLMeasure width;

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
                    if (FormatMap.TryGetValue(this.Format, out ImageFormat imageFormat) && imageFormat.Guid != bitmap.RawFormat.Guid)
                        throw new ArgumentException("The picture format in the stream and the parameter don't match");

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
            this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            this.Placement = XLPicturePlacement.MoveAndSize;
            this.Markers = new Dictionary<XLMarkerPosition, XLMarker>()
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

        public IXLCell BottomRightCell
        {
            get
            {
                return Markers[XLMarkerPosition.BottomRight].Cell;
            }

            private set
            {
                if (!value.Worksheet.Equals(this.Worksheet))
                    throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

                this.Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
            }
        }

        public XLPictureFormat Format { get; private set; }

        public IXLMeasure Height
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
                if ((Worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
                    throw new ArgumentException($"The picture ID '{value}' already exists.");

                id = value;
            }
        }

        public MemoryStream ImageStream { get; private set; }

        public IXLMeasure Left
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.X ?? XLMeasure.Zero; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), value, this.Top);
            }
        }

        public String Name
        {
            get { return name; }
            set
            {
                if (name == value) return;

                if ((Worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? this) != this)
                    throw new ArgumentException($"The picture name '{value}' already exists.");

                SetName(value);
            }
        }

        public IXLMeasure OriginalHeight { get; private set; }

        public IXLMeasure OriginalWidth { get; private set; }

        public XLPicturePlacement Placement { get; set; }

        public IXLMeasure Top
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.Y ?? XLMeasure.Zero; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the top offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), this.Left, value);
            }
        }

        public IXLCell TopLeftCell
        {
            get
            {
                return Markers[XLMarkerPosition.TopLeft].Cell;
            }

            private set
            {
                if (!value.Worksheet.Equals(this.Worksheet))
                    throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

                this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
            }
        }

        public IXLMeasure Width
        {
            get { return width; }
            set
            {
                if (this.Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the width, the placement should be FreeFloating or Move");
                width = value;
            }
        }

        public IXLWorksheet Worksheet { get; }

        internal IDictionary<XLMarkerPosition, XLMarker> Markers { get; private set; }

        internal String RelId { get; set; }

        /// <summary>
        /// Create a copy of the picture on a different worksheet.
        /// </summary>
        /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
        /// <returns>A created copy of the picture.</returns>
        public IXLPicture CopyTo(IXLWorksheet targetSheet)
        {
            return CopyTo((XLWorksheet)targetSheet);
        }

        public void Delete()
        {
            Worksheet.Pictures.Delete(this.Name);
        }

        // Used by Janitor.Fody
        private void DisposeManaged()
        {
            this.ImageStream.Dispose();
        }

#if _NET40_

        public void Dispose()
        {
            // net40 doesn't support Janitor.Fody, so let's dispose manually
            DisposeManaged();
        }

#else

        public void Dispose()
        {
            // Leave this empty (for non net40 targets) so that Janitor.Fody can do its work
        }

#endif

        /// <summary>
        /// Create a copy of the picture on the same worksheet.
        /// </summary>
        /// <returns>A created copy of the picture.</returns>
        public IXLPicture Duplicate()
        {
            return CopyTo(Worksheet);
        }

        public Tuple<IXLMeasure, IXLMeasure> GetOffset(XLMarkerPosition position)
        {
            var marker = Markers[position];
            if (marker == null)
                throw new NotSupportedException();

            return new Tuple<IXLMeasure, IXLMeasure>(marker.X, marker.Y);
        }

        public IXLPicture MoveTo(IXLMeasure left, IXLMeasure top)
        {
            this.Placement = XLPicturePlacement.FreeFloating;
            this.Left = left;
            this.Top = top;
            return this;
        }

        public IXLPicture MoveTo(IXLCell cell)
        {
            return MoveTo(cell, XLMeasure.Zero, XLMeasure.Zero);
        }

        public IXLPicture MoveTo(IXLCell cell, IXLMeasure xOffset, IXLMeasure yOffset)
        {
            this.Placement = XLPicturePlacement.Move;
            this.TopLeftCell = cell;
            var marker = this.Markers[XLMarkerPosition.TopLeft];
            marker.X = xOffset;
            marker.Y = yOffset;
            return this;
        }

        //public IXLPicture MoveTo(IXLCell cell, Point offset)
        //{
        //    if (cell == null) throw new ArgumentNullException(nameof(cell));
        //    return MoveTo(cell, new XLMeasure(offset.X, XLMeasureUnit.Pixels), new XLMeasure(offset.Y, XLMeasureUnit.Pixels));
        //}

        public IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell)
        {
            return MoveTo(fromCell, XLMeasure.Zero, XLMeasure.Zero, toCell, XLMeasure.Zero, XLMeasure.Zero);
        }

        //public IXLPicture MoveTo(IXLCell fromCell, IXLMeasure fromCellXOffset, IXLMeasure fromCellYOffset, IXLCell toCell, IXLMeasure toCellXOffset, IXLMeasure toCellYOffset)
        //{
        //    return MoveTo(fromCell, new Point(fromCellXOffset, fromCellYOffset), toCell, new Point(toCellXOffset, toCellYOffset));
        //}

        public IXLPicture MoveTo(IXLCell fromCell, IXLMeasure fromCellXOffset, IXLMeasure fromCellYOffset, IXLCell toCell, IXLMeasure toCellXOffset, IXLMeasure toCellYOffset)
        {
            if (fromCell == null) throw new ArgumentNullException(nameof(fromCell));
            if (toCell == null) throw new ArgumentNullException(nameof(toCell));
            this.Placement = XLPicturePlacement.MoveAndSize;

            this.TopLeftCell = fromCell;
            var marker = this.Markers[XLMarkerPosition.TopLeft];
            marker.X = fromCellXOffset;
            marker.Y = fromCellYOffset;

            this.BottomRightCell = toCell;
            marker = this.Markers[XLMarkerPosition.BottomRight];
            marker.X = toCellXOffset;
            marker.Y = toCellYOffset;

            return this;
        }

        public IXLPicture Scale(Double factor, Boolean relativeToOriginal = false)
        {
            return this.ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);
        }

        public IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false)
        {
            this.Height = new XLMeasure((relativeToOriginal ? this.OriginalHeight.Value : this.Height.Value) * factor, this.Height.Unit);
            return this;
        }

        public IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false)
        {
            this.Width = new XLMeasure((relativeToOriginal ? this.OriginalWidth.Value : this.Width.Value) * factor, this.Width.Unit);
            return this;
        }

        public IXLPicture WithPlacement(XLPicturePlacement value)
        {
            this.Placement = value;
            return this;
        }

        public IXLPicture WithSize(IXLMeasure width, IXLMeasure height)
        {
            this.Width = width;
            this.Height = height;
            return this;
        }

        internal IXLPicture CopyTo(XLWorksheet targetSheet)
        {
            if (targetSheet == null)
                targetSheet = Worksheet as XLWorksheet;

            IXLPicture newPicture;
            if (targetSheet == Worksheet)
                newPicture = targetSheet.AddPicture(ImageStream, Format);
            else
                newPicture = targetSheet.AddPicture(ImageStream, Format, Name);

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
                    newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), this.Left, this.Top);
                    break;

                case XLPicturePlacement.MoveAndSize:
                    var offset = GetOffset(XLMarkerPosition.BottomRight);
                    newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), this.Left, this.Top,
                                      targetSheet.Cell(BottomRightCell.Address), offset.Item1, offset.Item2);
                    break;
            }

            return newPicture;
        }

        internal void SetName(string value)
        {
            if (String.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Picture names cannot be empty");

            if (value.IndexOfAny(InvalidNameChars.ToCharArray()) != -1)
                throw new ArgumentException($"Picture names cannot contain any of the following characters: {InvalidNameChars}");

            if (value.Length > 31)
                throw new ArgumentException("Picture names cannot be more than 31 characters");

            name = value;
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
            this.OriginalWidth = new XLMeasure(bitmap.Width, XLMeasureUnit.Pixels);
            this.OriginalHeight = new XLMeasure(bitmap.Height, XLMeasureUnit.Pixels);

            this.width = OriginalWidth;
            this.height = OriginalHeight;
        }
    }
}
