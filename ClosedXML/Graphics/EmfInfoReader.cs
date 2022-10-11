using System.Drawing;
using System.IO;
using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// Metadata read of a vector EMF file. Specification: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/
    /// </summary>
    internal class EmfInfoReader : ImageInfoReader
    {
        private const uint EmfSignature = 0x464D4520; // ' EMF'

        protected override bool CheckHeader(Stream stream)
        {
            if (!stream.TryReadU32LE(out var type) || type != 0x1)
                return false;
            stream.Position += 36;
            if (!stream.TryReadU32LE(out var signature) || signature != EmfSignature)
                return false;
            stream.Position += 14;
            if (!stream.TryReadU16LE(out var reserved) || reserved != 0x0)
                return false;
            return true;
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            stream.Position += 24;
            var frame = ReadRectL(stream);
            return new XLPictureInfo(XLPictureFormat.Emf, Size.Empty, frame.Size);
        }

        private static Rectangle ReadRectL(Stream stream)
        {
            var left = stream.ReadS32LE();
            var top = stream.ReadS32LE();
            var right = stream.ReadS32LE();
            var bottom = stream.ReadS32LE();
            return new Rectangle(left, top, right - left, bottom - top);
        }
    }
}
