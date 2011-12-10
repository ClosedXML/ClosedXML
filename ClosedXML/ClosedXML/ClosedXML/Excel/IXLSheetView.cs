using System;

namespace ClosedXML.Excel
{
    public enum XLSheetViewOptions { Normal, PageBreakPreview, PageLayout }
    public interface IXLSheetView
    {
        /// <summary>
        /// Gets or sets the row after which the vertical split should take place.
        /// </summary>
        Int32 SplitRow { get; set; }
        /// <summary>
        /// Gets or sets the column after which the horizontal split should take place.
        /// </summary>
        Int32 SplitColumn { get; set; }
        //Boolean FreezePanes { get; set; }
        /// <summary>
        /// Freezes the top X rows.
        /// </summary>
        /// <param name="rows">The rows to freeze.</param>
        void FreezeRows(Int32 rows);
        /// <summary>
        /// Freezes the left X columns.
        /// </summary>
        /// <param name="columns">The columns to freeze.</param>
        void FreezeColumns(Int32 columns);
        /// <summary>
        /// Freezes the specified rows and columns.
        /// </summary>
        /// <param name="rows">The rows to freeze.</param>
        /// <param name="columns">The columns to freeze.</param>
        void Freeze(Int32 rows, Int32 columns);

        XLSheetViewOptions View { get; set; }

        IXLSheetView SetView(XLSheetViewOptions value);
    }
}
