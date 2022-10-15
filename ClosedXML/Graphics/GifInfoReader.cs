using System;
using System.Drawing;
using System.IO;
using ClosedXML.Excel.Drawings;
using ClosedXML.Utils;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// Read info about a GIF file. Gif file has no DPI, only pixel ratio.
    /// Specification: https://www.w3.org/Graphics/GIF/spec-gif89a.txt
    /// </summary>
    internal class GifInfoReader : ImageInfoReader
    {
        protected override bool CheckHeader(Stream stream)
        {
            Span<byte> s = stackalloc byte[6];
            if (stream.Read(s) != s.Length)
                return false;

            var hasSignature =
                s[0] == 0x47 && // 'G'
                s[1] == 0x49 && // 'I'
                s[2] == 0x46 && // 'F'
                s[3] == 0x38 && // '8'
                (s[4] == 0x37 || s[4] == 0x39) && // '7' or '9'
                s[5] == 0x61; // 'a'
            return hasSignature;
        }

        protected override XLPictureInfo ReadInfo(Stream stream)
        {
            stream.Position += 6; // header length
            var width = stream.ReadU16LE();
            var height = stream.ReadU16LE();
            return new XLPictureInfo(XLPictureFormat.Gif, new Size(width, height), Size.Empty);
        }
    }
}
