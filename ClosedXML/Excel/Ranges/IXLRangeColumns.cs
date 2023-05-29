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

        IXLStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these columns.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
