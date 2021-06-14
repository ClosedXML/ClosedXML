// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeRows : IEnumerable<IXLRangeRow>
    {
        IXLStyle Style { get; set; }

        /// <summary>
        /// Adds a row range to this group.
        /// </summary>
        /// <param name="rangeRow">The range row to add.</param>
        void Add(IXLRangeRow rangeRow);

        /// <summary>
        /// Returns the collection of cells.
        /// </summary>
        IXLCells Cells();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        IXLCells CellsUsed();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        /// <param name="includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCells CellsUsed(Boolean includeFormats);

        IXLCells CellsUsed(XLCellsUsedOptions options);

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

        /// <summary>
        /// Returns a set of range rows that are all touching or connected throughout in an unbroken sequence.
        /// </summary>
        IXLRangeRows Contiguous();

        /// <summary>
        /// Deletes all rows and shifts the rows below them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Returns the first row
        /// </summary>
        IXLRangeRow FirstRow();

        /// <summary>
        /// Returns the last row
        /// </summary>
        IXLRangeRow LastRow();

        void Select();

        IXLRangeRows SetDataType(XLDataType dataType);
    }
}
