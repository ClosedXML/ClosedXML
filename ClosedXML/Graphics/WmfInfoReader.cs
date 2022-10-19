using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;
using System;
using System.Drawing;
using System.IO;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// Reader of Windows Meta File.
    /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/4813e7fd-52d0-4f42-965f-228c8b7488d2
    /// http://formats.kaitai.io/wmf/index.html
    /// </summary>
    internal class WmfInfoReader : ImageInfoReader
    {
        protected override bool CheckHeader(Stream stream)
        {
            Span<byte> header = stackalloc byte[22];
            if (stream.Read(header) != header.Length)
                return false;

            var hasPlaceableHeader = header[0] == 0xD7 &&
                                     header[1] == 0xCD &&
                                     header[2] == 0xC6 &&
                                     header[3] == 0x9A;
            if (hasPlaceableHeader)
                return true;

            // File might not contain the placeable header (2.3.2.3), the header was added in a later revision.
            var type = GetU16LE(header, 0);
            var headerSize = GetU16LE(header, 2);
            var version = GetU16LE(header, 4);

            var hasWmfHeader =
                (type == MetafileType.MemoryMetaFile || type == MetafileType.DiskMetaFile) &&
                headerSize == 0x9 &&
                (version == 0x100 || version == 0x300);
            return hasWmfHeader;
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            Span<byte> header = stackalloc byte[22];
            stream.Read(header);
            var hasPlaceableHeader = header[0] == 0xD7;
            if (hasPlaceableHeader)
            {
                var placeable = new PlaceableHeader(header);
                if (placeable.CheckSum == placeable.CalculateCheckSum())
                {
                    // Excel ignores inch field of placeable header, but we don't
                    var widthHiMetric = ToHiMetric(placeable.BoundingBox.Width, placeable.Inch);
                    var heightHiMetric = ToHiMetric(placeable.BoundingBox.Height, placeable.Inch);
                    return new XLPictureInfo(XLPictureFormat.Wmf, Size.Empty, new Size(widthHiMetric, heightHiMetric));
                }
            }

            // Either no placeable header or it is corrupted. Skip header.
            stream.Position = 18;
            var complete = false;
            var viewportOrigin = Point.Empty;
            var viewportExtent = Point.Empty;
            while (!complete)
            {
                var recordSizeInWords = stream.ReadU32LE();
                var recordType = stream.ReadU16LE();
                switch (recordType)
                {
                    case RecordType.Eof:
                        complete = true;
                        break;

                    case RecordType.SetWindowExtent:
                        var viewportExtentY = stream.ReadS16LE();
                        var viewportExtentX = stream.ReadS16LE();
                        viewportExtent = new Point(viewportExtentX, viewportExtentY);
                        break;

                    case RecordType.SetWindowOrigin:
                        var viewportOriginY = stream.ReadS16LE();
                        var viewportOriginX = stream.ReadS16LE();
                        viewportOrigin = new Point(viewportOriginX, viewportOriginY);
                        break;

                    default:
                        stream.Position += (recordSizeInWords - 3) * 2;
                        break;
                }
            }

            if (viewportExtent == Point.Empty)
                throw new ArgumentException("Viewport extent is empty.");

            // Excel uses 96 logical units per inch
            var physSize = new Size(ToHiMetric(viewportExtent.X - viewportOrigin.X, 96), ToHiMetric(viewportExtent.Y - viewportOrigin.Y, 96));
            return new XLPictureInfo(XLPictureFormat.Wmf, Size.Empty, physSize);
        }

        private static uint GetU32LE(Span<byte> s, int i)
            => (uint)(s[i + 0] | (s[i + 1] << 8) | (s[i + 2] << 16) | (s[i + 3] << 24));

        private static ushort GetU16LE(Span<byte> s, int i)
            => (ushort)(s[i + 0] | (s[i + 1] << 8));

        private static short GetS16LE(Span<byte> s, int i)
            => (short)(s[i + 0] | (s[i + 1] << 8));

        private static int ToHiMetric(int size, double unitsPerInch) =>
            (int)Math.Round(size / unitsPerInch * 254d, MidpointRounding.AwayFromZero);

        private class PlaceableHeader
        {
            private readonly uint _key;
            private readonly ushort _resourceHandle;
            private readonly uint _reserved;
            public Rectangle BoundingBox { get; }
            public ushort Inch { get; }
            public ushort CheckSum { get; }

            public PlaceableHeader(Span<byte> s)
            {
                _key = GetU32LE(s, 0);
                _resourceHandle = GetU16LE(s, 4);
                var bboxLeft = GetS16LE(s, 6);
                var bboxTop = GetS16LE(s, 8);
                var bboxRight = GetS16LE(s, 10);
                var bboxBottom = GetS16LE(s, 12);
                BoundingBox = new Rectangle(bboxLeft, bboxTop, bboxRight - bboxLeft, bboxBottom - bboxTop);
                Inch = GetU16LE(s, 14);
                _reserved = GetU32LE(s, 16);
                CheckSum = GetU16LE(s, 20);
            }

            public ushort CalculateCheckSum()
            {
                ushort checkSum = 0;
                checkSum ^= (ushort)((_key & 0xFFFF0000) >> 16);
                checkSum ^= (ushort)(_key & 0xFFFF);
                checkSum ^= _resourceHandle;
                checkSum ^= (ushort)BoundingBox.Left;
                checkSum ^= (ushort)BoundingBox.Top;
                checkSum ^= (ushort)BoundingBox.Right;
                checkSum ^= (ushort)BoundingBox.Bottom;
                checkSum ^= Inch;
                checkSum ^= (ushort)((_reserved & 0xFFFF0000) >> 16);
                checkSum ^= (ushort)(_reserved & 0xFFFF);
                return checkSum;
            }
        }

        private static class MetafileType
        {
            public const ushort MemoryMetaFile = 0x001;
            public const ushort DiskMetaFile = 0x002;
        }

        private static class RecordType
        {
            public const ushort Eof = 0x0000;
            public const ushort SetWindowOrigin = 0x020B;
            public const ushort SetWindowExtent = 0x020C;
        }
    }
}
