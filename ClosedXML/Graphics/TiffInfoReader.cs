using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;
using System;
using System.IO;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A reader for baseline TIFF.
    /// Specification: https://www.itu.int/itudoc/itu-t/com16/tiff-fx/docs/tiff6.pdf
    /// </summary>
    internal class TiffInfoReader : ImageInfoReader
    {
        private delegate bool TryReadU16(Stream s, out ushort value);

        protected override bool CheckHeader(Stream stream)
        {
            if (!stream.TryReadU16BE(out var byteOrder))
                return false;

            var usesLittleEndian = byteOrder == ByteOrder.LittleEndian;
            var usesBigEndian = byteOrder == ByteOrder.BigEndian;
            if (!usesBigEndian && !usesLittleEndian)
                return false;

            TryReadU16 tryReadU16 = usesLittleEndian ? StreamExtensions.TryReadU16LE : StreamExtensions.TryReadU16BE;
            if (!tryReadU16(stream, out var version))
                return false;

            // The value (42) was chosen for its deep philosophical value :)
            return version == 42;
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            var byteOrder = stream.ReadU16BE();
            var usesLittleEndian = byteOrder == ByteOrder.LittleEndian;
            Func<Stream, uint> readU32 = usesLittleEndian ? StreamExtensions.ReadU32LE : StreamExtensions.ReadU32BE;
            Func<Stream, ushort> readU16 = usesLittleEndian ? StreamExtensions.ReadU16LE : StreamExtensions.ReadU16BE;

            stream.Position += 2; // skip version
            var ifdOffset = readU32(stream);
            stream.Position = ifdOffset;
            var entriesCount = readU16(stream);
            uint width = 0, height = 0, resolutionUnit = 2;
            double xResolution = 0, yResolution = 0;
            for (var i = 0; i < entriesCount; ++i)
            {
                var entryTag = readU16(stream);
                var entryType = readU16(stream);
                stream.Position += 4; // entryCount
                var nextEntryStart = stream.Position + 4;
                switch (entryTag)
                {
                    case Tag.ImageWidth:
                        width = ReadShortOrLone(stream, entryType, readU16, readU32);
                        break;

                    case Tag.ImageLength:
                        height = ReadShortOrLone(stream, entryType, readU16, readU32);
                        break;

                    case Tag.XResolution:
                        xResolution = ReadRational(stream, readU32);
                        break;

                    case Tag.YResolution:
                        yResolution = ReadRational(stream, readU32);
                        break;

                    case Tag.ResolutionUnit:
                        resolutionUnit = readU16(stream);
                        break;
                }
                stream.Position = nextEntryStart;
            }

            if (width == 0 || height == 0)
                throw new ArgumentException("Unable to determine dimensions of a TIFF.");

            var dpiX = ToDpi(xResolution, resolutionUnit);
            var dpiY = ToDpi(yResolution, resolutionUnit);
            return new XLPictureInfo(XLPictureFormat.Tiff, width, height, dpiX, dpiY);
        }

        private static uint ReadShortOrLone(Stream stream, ushort entryType, Func<Stream, ushort> readU16, Func<Stream, uint> readU32)
        {
            if (entryType == FieldType.Long)
                return readU32(stream);

            if (entryType == FieldType.Short)
                return readU16(stream);

            throw new ArgumentException("Expected only SHORT/LONG type.");
        }

        private static double ReadRational(Stream stream, Func<Stream, uint> readU32)
        {
            stream.Position = readU32(stream);
            var numerator = readU32(stream);
            var denominator = readU32(stream);
            return (double)numerator / denominator;
        }

        private static double ToDpi(double resolution, uint resolutionUnit) =>
            resolutionUnit switch
            {
                Unit.Inch => resolution,
                Unit.Cm => resolution * 2.54d,
                _ => 0
            };

        private static class ByteOrder
        {
            /// <summary>
            /// <c>II</c> like Intel in ASCII.
            /// </summary>
            public const ushort LittleEndian = 0x4949;

            /// <summary>
            /// <c>MM</c> like Motorola in ASCII.
            /// </summary>
            public const ushort BigEndian = 0x4D4D;
        }

        private static class Tag
        {
            public const ushort ImageWidth = 0x100;
            public const ushort ImageLength = 0x101;
            public const ushort XResolution = 0x11A;
            public const ushort YResolution = 0x11B;
            public const ushort ResolutionUnit = 0x128;
        }

        private static class FieldType
        {
            public const ushort Short = 3; // 2 bytes, unsigned
            public const ushort Long = 4; // 4 bytes, unsigned
        }

        private static class Unit
        {
            public const int Inch = 2;
            public const int Cm = 3;
        }
    }
}
