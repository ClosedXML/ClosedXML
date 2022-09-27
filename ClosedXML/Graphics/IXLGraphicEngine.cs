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
        /// Get the dimensions of a picture.
        /// </summary>
        /// <param name="imageStream">Stream is at the beginning of the image.</param>
        /// <param name="expectedFormat">The expected format of the image. Use <see cref="XLPictureFormat.Unknown"/> for auto detection.</param>
        /// <exception cref="ArgumentException">Unable to determine picture dimensions or format doesn't match the stream.</exception>
        XLPictureMetadata GetPictureMetadata(Stream imageStream, XLPictureFormat expectedFormat);
    }
}
