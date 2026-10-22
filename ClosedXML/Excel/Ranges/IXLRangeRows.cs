#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeRows : IEnumerable<IXLRangeRow>
    {
        /// <summary>
        /// Adds a row range to this group.
        /// </summary>
        /// <param name="rowRange">The row range to add.</param>
        void Add(IXLRangeRow rowRange);

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
        /// Deletes all rows and shifts the rows below them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjusts the height of all rows based on the contents of cells in these range rows.
        /// </summary>
        IXLRangeRows AdjustToContents();

        IXLRangeRows AdjustToContents(Int32 startColumn);

        IXLRangeRows AdjustToContents(Int32 startColumn, Int32 endColumn);

        IXLRangeRows AdjustToContents(Double minHeight, Double maxHeight);

        IXLRangeRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight);

        IXLRangeRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight);

        /// <summary>
        /// Returns a collection of the same type containing the rows that match the predicate.
        /// </summary>
        IXLRangeRows Where(Func<IXLRangeRow, Boolean> predicate);

        /// <summary>
        /// Returns a collection of the same type containing all rows after the first
        /// <paramref name="count"/> rows, in row-number order.
        /// </summary>
        IXLRangeRows Skip(Int32 count);

        /// <summary>
        /// Returns a collection of the same type containing the first <paramref name="count"/>
        /// rows, in row-number order.
        /// </summary>
        IXLRangeRows Take(Int32 count);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
