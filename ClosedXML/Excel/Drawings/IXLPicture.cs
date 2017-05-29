using System;
using System.Collections.Generic;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public enum XLPictureFormat
    {
        Bmp = 0,
        Gif = 1,
        Png = 2,
        Tiff = 3,
        Icon = 4,
        Pcx = 5,
        Jpeg = 6,
        Emf = 7,
        Wmf = 8
    }

    public interface IXLPicture : IDisposable
    {
        /// <summary>
        /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
        /// Default value is "jpeg"
        /// </summary>
        XLPictureFormat Format { get; }

        long Height { get; set; }
        MemoryStream ImageStream { get; }
        bool IsAbsolute { get; }
        long Left { get; set; }
        IList<IXLMarker> Markers { get; }
        String Name { get; set; }
        long Top { get; set; }
        long Width { get; set; }

        IXLPicture AtPosition(long left, long top);

        IXLPicture SetAbsolute();

        IXLPicture SetAbsolute(bool value);

        IXLMarker WithMarker(IXLMarker marker);

        IXLPicture WithSize(long width, long height);
    }
}
