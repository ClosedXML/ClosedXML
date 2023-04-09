using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRanges : IEnumerable<IXLRange>
    {
        /// <summary>
        /// Adds the specified range to this group.
        /// </summary>
        /// <param name="range">The range to add to this group.</param>
        void Add(IXLRangeBase range);

        void Add(IXLCell range);

        /// <summary>
        /// Removes the specified range from this group.
        /// </summary>
        /// <param name="range">The range to remove from this group.</param>
        bool Remove(IXLRange range);

        /// <summary>
        /// Removes ranges matching the criteria from the collection, optionally releasing their event handlers.
        /// </summary>
        /// <param name="match">Criteria to filter ranges. Only those ranges that satisfy the criteria will be removed.
        /// Null means the entire collection should be cleared.</param>
        /// <param name="releaseEventHandlers">Specify whether or not should removed ranges be unsubscribed from
        /// row/column shifting events. Until ranges are unsubscribed they cannot be collected by GC.</param>
        void RemoveAll(Predicate<IXLRange>? match = null, bool releaseEventHandlers = true);

        Int32 Count { get; }

        Boolean Contains(IXLRange range);

        /// <summary>
        /// Filter ranges from a collection that intersect the specified address. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        IEnumerable<IXLRange> GetIntersectedRanges(IXLRangeAddress rangeAddress);

        /// <summary>
        /// Filter ranges from a collection that intersect the specified address. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        IEnumerable<IXLRange> GetIntersectedRanges(IXLAddress address);

        /// <summary>
        /// Filter ranges from a collection that intersect the specified cell. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        IEnumerable<IXLRange> GetIntersectedRanges(IXLCell cell);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Creates a new data validation rule for the ranges collection, replacing the existing ones.
        /// </summary>
        IXLDataValidation CreateDataValidation();

        [Obsolete("Use CreateDataValidation() instead.")]
        IXLDataValidation SetDataValidation();

        /// <summary>
        /// Creates a named range out of these ranges.
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <para>The default scope for the named range is Workbook.</para>
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        IXLRanges AddToNamed(String rangeName);

        /// <summary>
        /// Creates a named range out of these ranges.
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// </summary>
        IXLRanges AddToNamed(String rangeName, XLScope scope);

        /// <summary>
        /// Creates a named range out of these ranges.
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// <param name="comment">The comments for the named range.</param>
        /// </summary>
        IXLRanges AddToNamed(String rangeName, XLScope scope, String comment);

        /// <summary>
        /// Sets the cells' value.
        /// <para>
        /// Setter will clear a formula, if the cell contains a formula.
        /// If the value is a text that starts with a single quote, setter will prefix the value with a single quote through
        /// <see cref="IXLStyle.IncludeQuotePrefix"/> in Excel too and the value of cell is set to to non-quoted text.
        /// </para>
        /// </summary>
        XLCellValue Value { set; }

        IXLRanges SetValue(XLCellValue value);

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
        /// Clears the contents of these ranges.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRanges Clear(XLClearOptions clearOptions = XLClearOptions.All);

        /// <summary>
        /// Create a new collection of ranges which are consolidated version of source ranges.
        /// </summary>
        IXLRanges Consolidate();

        void Select();
    }
}
