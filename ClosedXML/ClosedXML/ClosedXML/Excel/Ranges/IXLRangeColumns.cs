using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeColumns: IEnumerable<IXLRangeColumn>, IXLStylized
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
        /// Sets the formula for all cells in the columns in A1 notation.
        /// </summary>
        /// <value>
        /// The formula A1.
        /// </value>
        String FormulaA1 { set; }
        /// <summary>
        /// Sets the formula for all cells in the columns in R1C1 notation.
        /// </summary>
        /// <value>
        /// The formula R1C1.
        /// </value>
        String FormulaR1C1 { set; }

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
    }
}
