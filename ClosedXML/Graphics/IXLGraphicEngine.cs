using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// An interface to abstract away graphical functionality, like reading images, fonts or to determine a width of a text.
    /// </summary>
    public interface IXLGraphicEngine
    {
        /// <summary>
        /// Get the info about a picture (mostly dimensions).
        /// </summary>
        /// <param name="imageStream">Stream is at the beginning of the image.</param>
        /// <param name="expectedFormat">The expected format of the image. Use <see cref="XLPictureFormat.Unknown"/> for auto detection.</param>
        /// <exception cref="ArgumentException">Unable to determine picture dimensions or format doesn't match the stream.</exception>
        XLPictureInfo GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat);

        /// <summary>
        /// Get the height of a text with the font in pixels. Should be <c>EMHeight+descent</c>.
        /// </summary>
        double GetTextHeight(IXLFontBase font, double dpiY);

        /// <summary>
        /// Get the width of a text in pixels. Do not add any padding, there can be
        /// multiple spans of a texts with different fonts in a line.
        /// </summary>
        double GetTextWidth(string text, IXLFontBase font, double dpiX);

        /// <summary>
        /// The width of the widest 0-9 digit in pixels.
        /// </summary>
        /// <remarks>OOXML measures width of a column in multiples of widest 0-9 digit character in a normal style font.</remarks>
        double GetMaxDigitWidth(IXLFontBase font, double dpiX);

        /// <summary>
        /// Get font descent in pixels (positive value).
        /// </summary>
        /// <remarks>Excel is using OS/2 WinAscent/WinDescent for TrueType fonts (e.g. Calibri), not a correct font ascent/descent.</remarks>
        double GetDescent(IXLFontBase font, double dpiY);
    }
}
