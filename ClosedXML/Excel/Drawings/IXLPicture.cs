// Keep this file CodeMaid organised and cleaned
using System;
using System.Drawing;
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

        Int32 Height { get; set; }

        Int32 Id { get; }

        MemoryStream ImageStream { get; }

        Int32 Left { get; set; }

        String Name { get; set; }

        Int32 OriginalHeight { get; }

        Int32 OriginalWidth { get; }

        XLPicturePlacement Placement { get; set; }

        Int32 Top { get; set; }

        IXLCell TopLeftCell { get; }

        Int32 Width { get; set; }

        IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Create a copy of the picture on a different worksheet.
        /// </summary>
        /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
        /// <returns>A created copy of the picture.</returns>
        IXLPicture CopyTo(IXLWorksheet targetSheet);

        /// <summary>
        /// Deletes this picture.
        /// </summary>
        void Delete();

        /// <summary>
        /// Create a copy of the picture on the same worksheet.
        /// </summary>
        /// <returns>A created copy of the picture.</returns>
        IXLPicture Duplicate();

        Point GetOffset(XLMarkerPosition position);

        IXLPicture MoveTo(Int32 left, Int32 top);

        IXLPicture MoveTo(IXLCell cell);

        IXLPicture MoveTo(IXLCell cell, Int32 xOffset, Int32 yOffset);

        IXLPicture MoveTo(IXLCell cell, Point offset);

        IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell);

        IXLPicture MoveTo(IXLCell fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXLCell toCell, Int32 toCellXOffset, Int32 toCellYOffset);

        IXLPicture MoveTo(IXLCell fromCell, Point fromOffset, IXLCell toCell, Point toOffset);

        IXLPicture Scale(Double factor, Boolean relativeToOriginal = false);

        IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false);

        IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false);

        IXLPicture WithPlacement(XLPicturePlacement value);

        IXLPicture WithSize(Int32 width, Int32 height);
    }
}
