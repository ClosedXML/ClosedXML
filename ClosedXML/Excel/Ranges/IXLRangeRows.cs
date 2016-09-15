﻿using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeRows : IEnumerable<IXLRangeRow>, IDisposable
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
        /// <param name="includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCells CellsUsed(Boolean includeFormats);

        /// <summary>
        /// Deletes all rows and shifts the rows below them accordingly.
        /// </summary>
        void Delete();

        IXLStyle Style { get; set; }

        IXLRangeRows SetDataType(XLCellValues dataType);

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);

        void Select();
    }
}
