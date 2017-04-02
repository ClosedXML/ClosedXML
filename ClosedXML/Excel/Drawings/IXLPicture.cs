using System;
using System.Drawing;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLPicture : IDisposable
    {
        IXLAddress BottomRightCellAddress { get; }

        /// <summary>
        /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
        /// Default value is "jpeg"
        /// </summary>
        XLPictureFormat Format { get; }

        Int32 Height { get; set; }

        MemoryStream ImageStream { get; }

        Int32 Left { get; set; }

        String Name { get; set; }

        Int32 OriginalHeight { get; }

        Int32 OriginalWidth { get; }

        XLPicturePlacement Placement { get; set; }

        Int32 Top { get; set; }

        IXLAddress TopLeftCellAddress { get; }

        Int32 Width { get; set; }

        IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Deletes this picture.
        /// </summary>
        void Delete();

        Point GetOffset(XLMarkerPosition position);

        IXLPicture MoveTo(Int32 left, Int32 top);

        IXLPicture MoveTo(IXLAddress cell);

        IXLPicture MoveTo(IXLAddress cell, Int32 xOffset, Int32 yOffset);

        IXLPicture MoveTo(IXLAddress cell, Point offset);

        IXLPicture MoveTo(IXLAddress fromCell, IXLAddress toCell);

        IXLPicture MoveTo(IXLAddress fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXLAddress toCell, Int32 toCellXOffset, Int32 toCellYOffset);

        IXLPicture MoveTo(IXLAddress fromCell, Point fromOffset, IXLAddress toCell, Point toOffset);

        IXLPicture Scale(Double factor, Boolean relativeToOriginal = false);

        IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false);

        IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false);

        IXLPicture WithPlacement(XLPicturePlacement value);

        IXLPicture WithSize(Int32 width, Int32 height);
    }
}
