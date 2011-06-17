using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLTableRows: IEnumerable<IXLTableRow>
    {
        /// <summary>
        /// Adds a table row to this group.
        /// </summary>
        /// <param name="tableRow">The row table to add.</param>
        void Add(IXLTableRow tableRow);

        /// <summary>
        /// Clears the contents of the rows (including styles).
        /// </summary>
        void Clear();

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
        /// <param name="includeStyles">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCells CellsUsed(Boolean includeStyles);

        IXLStyle Style { get; set; }

        
    }
}
