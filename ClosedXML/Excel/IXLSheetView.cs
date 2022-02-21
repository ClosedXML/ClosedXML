// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public enum XLSheetViewOptions { Normal, PageBreakPreview, PageLayout }

    public interface IXLSheetView
    {
        /// <summary>
        /// Gets or sets the column after which the horizontal split should take place.
        /// </summary>
        Int32 SplitColumn { get; set; }

        /// <summary>
        /// Gets or sets the row after which the vertical split should take place.
        /// </summary>
        Int32 SplitRow { get; set; }

        /// <summary>
        /// Gets or sets the location of the top left visible cell
        /// </summary>
        /// <value>
        /// The scroll position's top left cell.
        /// </value>
        IXLAddress TopLeftCellAddress { get; set; }

        XLSheetViewOptions View { get; set; }

        IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Window zoom magnification for current view representing percent values. Horizontal and vertical scale together.
        /// </summary>
        /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
        Int32 ZoomScale { get; set; }

        /// <summary>
        /// Zoom magnification to use when in normal view. Horizontal and vertical scale together
        /// </summary>
        /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
        Int32 ZoomScaleNormal { get; set; }

        /// <summary>
        /// Zoom magnification to use when in page layout view. Horizontal and vertical scale together.
        /// </summary>
        /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
        Int32 ZoomScalePageLayoutView { get; set; }

        /// <summary>
        /// Zoom magnification to use when in page break preview. Horizontal and vertical scale together.
        /// </summary>
        /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
        Int32 ZoomScaleSheetLayoutView { get; set; }

        /// <summary>
        /// Freezes the specified rows and columns.
        /// </summary>
        /// <param name="rows">The rows to freeze.</param>
        /// <param name="columns">The columns to freeze.</param>
        void Freeze(Int32 rows, Int32 columns);

        /// <summary>
        /// Freezes the left X columns.
        /// </summary>
        /// <param name="columns">The columns to freeze.</param>
        void FreezeColumns(Int32 columns);

        //Boolean FreezePanes { get; set; }
        /// <summary>
        /// Freezes the top X rows.
        /// </summary>
        /// <param name="rows">The rows to freeze.</param>
        void FreezeRows(Int32 rows);

        IXLSheetView SetView(XLSheetViewOptions value);
    }
}
