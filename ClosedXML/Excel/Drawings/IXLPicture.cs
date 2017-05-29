using System;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLPicture : IDisposable
    {
        IXLCell BottomRightCell { get; }

        /// <summary>
        /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
        /// Default value is "jpeg"
        /// </summary>
        XLPictureFormat Format { get; }

        long Height { get; set; }
        MemoryStream ImageStream { get; }
        long Left { get; set; }
        String Name { get; set; }
        long OriginalHeight { get; }
        long OriginalWidth { get; }
        XLPicturePlacement Placement { get; set; }
        long Top { get; set; }
        IXLCell TopLeftCell { get; }
        long Width { get; set; }

        IXLPicture AtPosition(long left, long top);

        IXLPicture AtPosition(IXLCell cell);

        IXLPicture AtPosition(IXLCell fromCell, IXLCell toCell);

        void ScaleHeight(Double factor, Boolean relativeToOriginal = false);

        void ScaleWidth(Double factor, Boolean relativeToOriginal = false);

        IXLPicture WithPlacement(XLPicturePlacement value);

        IXLPicture WithSize(long width, long height);
    }
}
