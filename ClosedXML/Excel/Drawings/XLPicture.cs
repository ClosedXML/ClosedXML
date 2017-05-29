using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Excel.Drawings
{
    internal class XLPicture : IXLPicture
    {
        private readonly IXLWorksheet _worksheet;

        private XLPicture(IXLWorksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            this._worksheet = worksheet;
            this.Placement = XLPicturePlacement.MoveAndSize;
        }

        internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
            : this(worksheet)
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

            var formats = typeof(ImageFormat).GetProperties(BindingFlags.Static | BindingFlags.Public)
                .Where(pi => (pi.GetValue(null, null) as ImageFormat).Guid.Equals(bitmap.RawFormat.Guid));

            if (!formats.Any() || formats.Count() > 1)
                throw new ArgumentException("Unsupported or unknown image format in bitmap");

            this.Format = Enum.Parse(typeof(XLPictureFormat), formats.Single().Name, true).CastTo<XLPictureFormat>();
        }

        public IXLCell BottomRightCell
        {
            get
            {
                if (this.Markers.Count > 1)
                {
                    var marker = this.Markers.Skip(1).First();
                    return _worksheet.Cell(marker.RowId, marker.ColumnId);
                }
                else
                    return null;
            }

            private set
            {
                while (this.Markers.Count > 1)
                {
                    this.Markers.RemoveAt(this.Markers.Count - 1);
                }

                this.Markers.Add(new XLMarker()
                {
                    ColumnId = value.WorksheetColumn().ColumnNumber(),
                    RowId = value.WorksheetRow().RowNumber()
                });
            }
        }

        public XLPictureFormat Format { get; private set; }
        private long height;

        public long Height
        {
            get
            {
                return height;
            }
            set
            {
                if (this.Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the height, the placement should be FreeFloating or Move");
                height = value;
            }
        }

        public MemoryStream ImageStream { get; private set; }

        public long left;

        public long Left
        {
            get { return left; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");
                left = value;
            }
        }

        public IList<IXLMarker> Markers { get; private set; } = new List<IXLMarker>();
        public String Name { get; set; }
        public long OriginalHeight { get; private set; }

        public long OriginalWidth { get; private set; }

        public XLPicturePlacement Placement { get; set; }
        private long top;

        public long Top
        {
            get { return top; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the top offset, the placement should be FreeFloating");
                top = value;
            }
        }

        public IXLCell TopLeftCell
        {
            get
            {
                if (this.Markers.Any())
                {
                    var marker = this.Markers.First();
                    return _worksheet.Cell(marker.RowId, marker.ColumnId);
                }
                else
                    return null;
            }

            private set
            {
                if (this.Markers.Any())
                    this.Markers.RemoveAt(0);

                this.Markers.Insert(0, new XLMarker()
                {
                    ColumnId = value.WorksheetColumn().ColumnNumber(),
                    RowId = value.WorksheetRow().RowNumber()
                });
            }
        }

        private long width;

        public long Width
        {
            get
            {
                return width;
            }
            set
            {
                if (this.Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the width, the placement should be FreeFloating or Move");
                width = value;
            }
        }

        public IXLPicture AtPosition(long left, long top)
        {
            this.Placement = XLPicturePlacement.FreeFloating;
            this.Left = left;
            this.Top = top;
            return this;
        }

        public IXLPicture AtPosition(IXLCell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            this.Placement = XLPicturePlacement.Move;
            this.TopLeftCell = cell;
            return this;
        }

        public IXLPicture AtPosition(IXLCell fromCell, IXLCell toCell)
        {
            if (fromCell == null) throw new ArgumentNullException(nameof(fromCell));
            this.Placement = XLPicturePlacement.MoveAndSize;
            this.TopLeftCell = fromCell;

            if (toCell != null)
                this.BottomRightCell = toCell;

            return this;
        }

        public void Dispose()
        {
            this.ImageStream.Dispose();
        }

        public void ScaleHeight(Double factor, Boolean relativeToOriginal = false)
        {
            this.Height = Convert.ToInt64((relativeToOriginal ? this.OriginalHeight : this.Height) * factor);
        }

        public void ScaleWidth(Double factor, Boolean relativeToOriginal = false)
        {
            this.Width = Convert.ToInt64((relativeToOriginal ? this.OriginalWidth : this.Width) * factor);
        }

        public IXLPicture WithPlacement(XLPicturePlacement value)
        {
            this.Placement = value;
            return this;
        }

        public IXLPicture WithSize(long width, long height)
        {
            this.Width = width;
            this.Height = height;
            return this;
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
