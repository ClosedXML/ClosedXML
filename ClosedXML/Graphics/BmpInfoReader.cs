using System;
using System.Drawing;
using System.IO;
using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A reader for BMP for Windows and OS/2.
    /// Specification:
    /// https://www.fileformat.info/format/bmp/corion.htm
    /// https://www.fileformat.info/format/bmp/egff.htm
    /// https://www.fileformat.info/format/os2bmp/egff.htm
    /// </summary>
    internal class BmpInfoReader : ImageInfoReader
    {
        protected override bool CheckHeader(Stream stream)
        {
            Span<byte> s = stackalloc byte[2];
            if (stream.Read(s) != s.Length)
                return false;

            // Excel can't read V1.x CI-Color Icon, CP-Color Pointer, IC-Icon or PT-Pointer for OS/2, so don't decode them.
            return s[0] == 'B' && (s[1] == 'M' || s[0] == 'A');
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            stream.Position += 14;
            var infoHeaderSize = stream.ReadS32LE();
            // BMP Version 1.x, used by IBM OS/2 1.x and Win 2.0 and later
            if (infoHeaderSize == 12)
                return ReadBmpV1X(stream);

            // BMP Version 2.x used by IBM OS/2 has a different overall structure, but width/height and resolution have same offsets as V3.x
            // BMP Version 3.x has dimension and resolution at same offsets and V4.x+ only add fields
            return ReadBmpV2X(stream);
        }

        private static XLPictureInfo ReadBmpV1X(Stream stream)
        {
            var widthPx = stream.ReadU16LE();
            var heightPx = stream.ReadU16LE();
            return new XLPictureInfo(XLPictureFormat.Bmp, new Size(widthPx, heightPx), Size.Empty);
        }

        private static XLPictureInfo ReadBmpV2X(Stream stream)
        {
            var widthPx = stream.ReadU32LE();
            var heightPx = stream.ReadU32LE();
            stream.Position += 12;
            var dpiX = PixelsPerMeterToDpi(stream.ReadU32LE());
            var dpiY = PixelsPerMeterToDpi(stream.ReadU32LE());
            return new XLPictureInfo(XLPictureFormat.Bmp, widthPx, heightPx, dpiX, dpiY);
        }

        private static double PixelsPerMeterToDpi(uint pixelsPerMeter)
            => pixelsPerMeter * 2.54d / 100d;
    }
}
