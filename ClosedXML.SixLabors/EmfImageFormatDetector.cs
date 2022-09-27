using SixLabors.ImageSharp.Formats;
using System;

namespace ClosedXML.Graphics
{
    internal class EmfImageFormatDetector : IImageFormatDetector
    {
        public int HeaderSize => 44;

        public IImageFormat DetectFormat(ReadOnlySpan<byte> header)
        {
            return header.Length >= HeaderSize && IsEmf(header) ? EmfFormat.Instance : null;
        }

        private static bool IsEmf(ReadOnlySpan<byte> header)
        {
            var versionIsOne =
                header[0] == 0x1 &&
                header[1] == 0x0 &&
                header[2] == 0x0 &&
                header[3] == 0x0;
            if (!versionIsOne)
                return false;

            var signatureIsEmf =
                header[40] == 0x20 && // ' '
                header[41] == 0x45 && // 'E'
                header[42] == 0x4D && // 'M'
                header[43] == 0x46;   // 'F'
            return signatureIsEmf;
        }
    }
}
