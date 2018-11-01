// Keep this file CodeMaid organised and cleaned
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

        IXLMeasure Height { get; set; }

        Int32 Id { get; }

        MemoryStream ImageStream { get; }

        IXLMeasure Left { get; set; }

        String Name { get; set; }

        IXLMeasure OriginalHeight { get; }

        IXLMeasure OriginalWidth { get; }

        XLPicturePlacement Placement { get; set; }

        IXLMeasure Top { get; set; }

        IXLCell TopLeftCell { get; }

        IXLMeasure Width { get; set; }

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

        Tuple<IXLMeasure, IXLMeasure> GetOffset(XLMarkerPosition position);

        IXLPicture MoveTo(IXLMeasure left, IXLMeasure top);

        IXLPicture MoveTo(IXLCell cell);

        IXLPicture MoveTo(IXLCell cell, IXLMeasure xOffset, IXLMeasure yOffset);

        IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell);

        IXLPicture MoveTo(IXLCell fromCell, IXLMeasure fromCellXOffset, IXLMeasure fromCellYOffset, IXLCell toCell, IXLMeasure toCellXOffset, IXLMeasure toCellYOffset);

        IXLPicture Scale(Double factor, Boolean relativeToOriginal = false);

        IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false);

        IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false);

        IXLPicture WithPlacement(XLPicturePlacement value);

        IXLPicture WithSize(IXLMeasure width, IXLMeasure height);
    }
}
