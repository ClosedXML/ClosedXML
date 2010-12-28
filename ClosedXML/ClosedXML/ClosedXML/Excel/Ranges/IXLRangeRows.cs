using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeRows: IEnumerable<IXLRangeRow>, IXLStylized
    {
        /// <summary>
        /// Adds a row range to this group.
        /// </summary>
        /// <param name="rowRange">The row range to add.</param>
        void Add(IXLRangeRow rowRange);
        /// <summary>
        /// Clears the contents of the rows (including styles).
        /// </summary>
        void Clear();
        /// <summary>
        /// Sets the formula for all cells in the rows in A1 notation.
        /// </summary>
        /// <value>
        /// The formula A1.
        /// </value>
        String FormulaA1 { set; }
        /// <summary>
        /// Sets the formula for all cells in the rows in R1C1 notation.
        /// </summary>
        /// <value>
        /// The formula R1C1.
        /// </value>
        String FormulaR1C1 { set; }

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
    }
}
