// Keep this file CodeMaid organised and cleaned
using SkiaSharp;
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

        int Height { get; set; }

        int Id { get; }

        MemoryStream ImageStream { get; }

        float Left { get; set; }

        string Name { get; set; }

        int OriginalHeight { get; }

        int OriginalWidth { get; }

        XLPicturePlacement Placement { get; set; }

        float Top { get; set; }

        IXLCell TopLeftCell { get; }

        int Width { get; set; }

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

        SKPoint GetOffset(XLMarkerPosition position);

        IXLPicture MoveTo(float left, float top);

        IXLPicture MoveTo(IXLCell cell);

        IXLPicture MoveTo(IXLCell cell, int xOffset, int yOffset);

        IXLPicture MoveTo(IXLCell cell, SKPoint offset);

        IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell);

        IXLPicture MoveTo(IXLCell fromCell, int fromCellXOffset, int fromCellYOffset, IXLCell toCell, int toCellXOffset, int toCellYOffset);

        IXLPicture MoveTo(IXLCell fromCell, SKPoint fromOffset, IXLCell toCell, SKPoint toOffset);

        IXLPicture Scale(double factor, bool relativeToOriginal = false);

        IXLPicture ScaleHeight(double factor, bool relativeToOriginal = false);

        IXLPicture ScaleWidth(double factor, bool relativeToOriginal = false);

        IXLPicture WithPlacement(XLPicturePlacement value);

        IXLPicture WithSize(int width, int height);
    }
}
