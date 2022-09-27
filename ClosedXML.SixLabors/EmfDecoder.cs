using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.PixelFormats;
using System;
using System.IO;
using System.Threading;

namespace ClosedXML.Graphics
{
    internal class EmfDecoder : IImageDecoder, IImageInfoDetector
    {
        public IImageInfo Identify(Configuration configuration, Stream stream, CancellationToken cancellationToken)
        {
            stream.Position = 8;
            var bounds = ReadRectL(stream);
            var frame = ReadRectL(stream);
            var imageInfo = new ImageInfo(bounds.Width + 1, bounds.Height + 1);
            var metadata = imageInfo.Metadata.GetFormatMetadata(EmfFormat.Instance);
            metadata.Frame = frame;
            return imageInfo;
        }

        public Image<TPixel> Decode<TPixel>(Configuration configuration, Stream stream, CancellationToken cancellationToken) where TPixel : unmanaged, IPixel<TPixel>
            => throw new NotSupportedException("Decoder can be used only for Identity.");

        public Image Decode(Configuration configuration, Stream stream, CancellationToken cancellationToken)
            => throw new NotSupportedException("Decoder can be used only for Identity.");

        private static Rectangle ReadRectL(Stream stream)
        {
            var left = ReadInt32LittleEndian(stream);
            var top = ReadInt32LittleEndian(stream);
            var right = ReadInt32LittleEndian(stream);
            var bottom = ReadInt32LittleEndian(stream);
            return new Rectangle(left, top, right - left, bottom - top);
        }

        private static int ReadInt32LittleEndian(Stream stream)
        {
            var b1 = ReadByte(stream);
            var b2 = ReadByte(stream);
            var b3 = ReadByte(stream);
            var b4 = ReadByte(stream);
            return b4 << 24 | b3 << 16 | b2 << 8 | b1;
        }

        private static byte ReadByte(Stream stream)
        {
            var b = stream.ReadByte();
            if (b == -1)
                throw new InvalidImageContentException("Unexpected end of stream.");
            return (byte)b;
        }
    }
}
