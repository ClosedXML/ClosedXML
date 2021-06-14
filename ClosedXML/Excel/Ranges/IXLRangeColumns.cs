// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeColumns : IEnumerable<IXLRangeColumn>
    {
        IXLStyle Style { get; set; }

        /// <summary>
        /// Adds a column range to this group.
        /// </summary>
        /// <param name="rangeColumn">The column range to add.</param>
        void Add(IXLRangeColumn rangeColumn);

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
        /// Clears the contents of these columns.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All);

        /// <summary>
        /// Returns a set of range columns that are all touching or connected throughout in an unbroken sequence.
        /// </summary>
        IXLRangeColumns Contiguous();

        /// <summary>
        /// Deletes all columns and shifts the columns at the right of them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Returns the first column
        /// </summary>
        IXLRangeColumn FirstColumn();

        /// <summary>
        /// Returns the last column
        /// </summary>
        IXLRangeColumn LastColumn();

        void Select();

        IXLRangeColumns SetDataType(XLDataType dataType);
    }
}
