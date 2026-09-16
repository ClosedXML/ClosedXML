#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeColumns : IEnumerable<IXLRangeColumn>
    {
        /// <summary>
        /// Adds a column range to this group.
        /// </summary>
        /// <param name="columRange">The column range to add.</param>
        void Add(IXLRangeColumn columRange);

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
        /// <param name="options">The options to determine whether a cell is used.</param>
        IXLCells CellsUsed(XLCellsUsedOptions options);

        /// <summary>
        /// Deletes all columns and shifts the columns at the right of them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjusts the width of all columns based on the contents of cells in these range columns.
        /// </summary>
        IXLRangeColumns AdjustToContents();

        IXLRangeColumns AdjustToContents(Int32 startRow);

        IXLRangeColumns AdjustToContents(Int32 startRow, Int32 endRow);

        IXLRangeColumns AdjustToContents(Double minWidth, Double maxWidth);

        IXLRangeColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth);

        IXLRangeColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth);

        /// <summary>
        /// Returns a collection of the same type containing the columns that match the predicate.
        /// </summary>
        IXLRangeColumns Where(Func<IXLRangeColumn, Boolean> predicate);

        /// <summary>
        /// Returns a collection of the same type containing all columns after the first
        /// <paramref name="count"/> columns, in column-number order.
        /// </summary>
        IXLRangeColumns Skip(Int32 count);

        /// <summary>
        /// Returns a collection of the same type containing the first <paramref name="count"/>
        /// columns, in column-number order.
        /// </summary>
        IXLRangeColumns Take(Int32 count);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these columns.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
