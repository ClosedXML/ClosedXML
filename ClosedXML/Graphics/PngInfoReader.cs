using System;
using System.IO;
using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;

namespace ClosedXML.Graphics
{
    internal class PngInfoReader : ImageInfoReader
    {
        private const int CrcLength = 4;
        private const int SkippedHeaderLength = 5;

        private int[] MagicBytes { get; } = { 137, 80, 78, 71, 13, 10, 26, 10 };

        private const int HeaderType = 0x49484452; // IHDR
        private const int PhysicalDimensionType = 0x70485973; // pHYs

        protected override bool CheckHeader(Stream stream)
        {
            foreach (var magicByte in MagicBytes)
            {
                var streamByte = stream.ReadByte();
                if (streamByte != magicByte || streamByte == -1)
                    return false;
            }
            return true;
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            stream.Position += MagicBytes.Length;
            var hdrLength = stream.ReadU32BE();
            if (hdrLength != 13)
                throw CorruptedException("Header length must be 13.");
            if (ReadType(stream) != HeaderType)
                throw CorruptedException("First chunk type must be IHDR.");

            var width = stream.ReadU32BE();
            var height = stream.ReadU32BE();

            stream.Position += SkippedHeaderLength + CrcLength;

            uint pixelsPerUnitX = 0, pixelsPerUnitY = 0;
            while (stream.TryReadU32BE(out var chunkLength))
            {
                var chunkType = ReadType(stream);
                if (chunkType == PhysicalDimensionType)
                {
                    pixelsPerUnitX = stream.ReadU32BE();
                    pixelsPerUnitY = stream.ReadU32BE();
                    var unit = stream.ReadU8();
                    var isUnitMeter = unit == 1;
                    if (!isUnitMeter)
                        pixelsPerUnitX = pixelsPerUnitY = 0;

                    break;
                }

                stream.Position += chunkLength + CrcLength;
            }

            var dpiX = PixelsPerMeterToDpi(pixelsPerUnitX);
            var dpiY = PixelsPerMeterToDpi(pixelsPerUnitY);
            return new XLPictureInfo(XLPictureFormat.Png, width, height, dpiX, dpiY);
        }

        private static uint ReadType(Stream stream) => stream.ReadU32BE();

        private static ArgumentException CorruptedException(string text) => new($"PNG is corrupted. {text}");
        
        private static double PixelsPerMeterToDpi(uint ppm)
        {
            // Conversion from the common integer dots-per-inch to pixels-per-meter is lossy, so instead of 96 we get 95.9866
            return ppm * 0.0254d;
        }
    }
}
