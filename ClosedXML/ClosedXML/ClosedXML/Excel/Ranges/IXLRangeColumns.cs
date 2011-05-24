using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeColumns: IEnumerable<IXLRangeColumn>
    {
        /// <summary>
        /// Clears the contents of the columns (including styles).
        /// </summary>
        void Clear();

        /// <summary>
        /// Adds a column range to this group.
        /// </summary>
        /// <param name="columRange">The column range to add.</param>
        void Add(IXLRangeColumn columRange);

        /// <summary>
        /// Returns the collection of cells in this column.
        /// </summary>
        IXLCells Cells();

        /// <summary>
        /// Returns the collection of cells that have a value in this column.
        /// </summary>
        IXLCells CellsUsed();

        /// <summary>
        /// Returns the collection of cells that have a value in this column.
        /// </summary>
        /// <param name="includeStyles">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCells CellsUsed(Boolean includeStyles);

        /// <summary>
        /// Deletes all columns and shifts the columns at the right of them accordingly.
        /// </summary>
        void Delete();

        IXLStyle Style { get; set; }

        

    }
}
