﻿using System;
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
        IXLCells CellsUsed(Boolean includeFormats);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);

        void Select();
    }
}
