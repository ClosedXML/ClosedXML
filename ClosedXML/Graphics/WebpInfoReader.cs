#nullable disable

using ClosedXML.Utils;
using System;
using System.Drawing;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// Reader of dimensions for WebP image format.
    /// </summary>
    internal class WebpInfoReader : ImageInfoReader
    {
        private const int Vp8ChunkMagicBytes = 0x9d012a;
        private const int Vp8LChunkMagicByte = 0x2F;

        private static readonly UInt32 LossyVp8Code = "VP8 ".ToMagicNumber();
        private static readonly UInt32 LosslessVp8Code = "VP8L".ToMagicNumber();
        private static readonly UInt32 ExtendedV8Code = "VP8X".ToMagicNumber();

        protected override bool CheckHeader(Stream stream)
        {
            Span<byte> header = stackalloc byte[12];
            if (stream.Read(header) != header.Length)
            {
                return false;
            }

            return header[0] == 'R' &&
                   header[1] == 'I' &&
                   header[2] == 'F' &&
                   header[3] == 'F' &&
                   header[8] == 'W' &&
                   header[9] == 'E' &&
                   header[10] == 'B' &&
                   header[11] == 'P';
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            // Skip header and file size
            stream.Position += 12;

            var chunkCode = stream.ReadU32BE();

            // Skip chunk size
            stream.Position += 4;
            if (chunkCode == ExtendedV8Code)
            {
                // https://developers.google.com/speed/webp/docs/riff_container#extended_file_format
                // Skip image features
                stream.Position += 4;
                var width = stream.ReadU24LE() + 1;
                var height = stream.ReadU24LE() + 1;

                // There is a potential EXIF/XMP chunk in extended format, but use default DPI to keep it simple.
                return new XLPictureInfo(XLPictureFormat.Webp, new Size(width, height), Size.Empty, 72, 72);
            }

            if (chunkCode == LossyVp8Code)
            {
                // https://datatracker.ietf.org/doc/html/rfc6386#section-9.1
                // First 3 bytes are a frame tag. It's read as a big endian for easier processing
                var frameTag = stream.ReadU24LE();
                var isKeyFrame = (frameTag & 1) == 0;
                if (!isKeyFrame)
                {
                    throw new ArgumentException("Image is not a key frame.");
                }

                var showFrameFlag = ((frameTag >> 4) & 0x1) == 1;
                if (!showFrameFlag)
                {
                    throw new ArgumentException("Frame is not visible.");
                }

                // Next 3 bytes are magic bytes
                var magicBytes = stream.ReadU24BE();
                if (magicBytes != Vp8ChunkMagicBytes)
                {
                    throw new ArgumentException("Invalid magic bytes for VP8 lossy chunk.");
                }

                // Scaling is used only for rendering, underlaying data are unscaled
                var widthAndScale = stream.ReadU16LE();
                var width = GetSize(widthAndScale);

                var heightAndScale = stream.ReadU16LE();
                var height = GetSize(heightAndScale);

                return new XLPictureInfo(XLPictureFormat.Webp, new Size(width, height), Size.Empty, 72, 72);

                static int GetSize(ushort sizeAndScale)
                {
                    var size = sizeAndScale & 0x3FFF;
                    var scale = sizeAndScale >> 14;
                    return scale switch
                    {
                        0 => size,
                        1 => size * 5 / 4,
                        2 => size * 5 / 3,
                        _ => size * 2
                    };
                }
            }

            if (chunkCode == LosslessVp8Code)
            {
                // https://developers.google.com/speed/webp/docs/webp_lossless_bitstream_specification
                var magic = stream.ReadByte();
                if (magic != Vp8LChunkMagicByte)
                {
                    throw new ArgumentException("Invalid magic for VP8L chunk.");
                }

                Span<byte> header = stackalloc byte[4];
                var readBytes = stream.Read(header);
                if (readBytes != 4)
                {
                    throw new ArgumentException("Unexpected end of file.");
                }

                // Width is 14 bits and height is 14 bit, packed into 4 bytes
                var width = header[0] + ((header[1] & 0x3F) << 8) + 1;
                var height = ((header[1] & 0xC0) >> 6) + (header[2] << 2) + ((header[3] & 0xF) << 10) + 1;

                return new XLPictureInfo(XLPictureFormat.Webp, new Size(width, height), Size.Empty, 72, 72);
            }

            throw new ArgumentException("Invalid chunk for WebP file.");
        }
    }
}
