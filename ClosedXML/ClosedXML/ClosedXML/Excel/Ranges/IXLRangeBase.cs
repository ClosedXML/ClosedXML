using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLScope { Workbook, Worksheet };

    public interface IXLRangeBase: IXLStylized
    {
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

        /// <summary>
        /// Gets an object with the boundaries of this range.
        /// </summary>
        IXLRangeAddress RangeAddress { get; }
        /// <summary>
        /// Returns the first cell of this range.
        /// </summary>
        IXLCell FirstCell();
        /// <summary>
        /// Returns the first cell with a value of this range.
        /// <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
        /// </summary>
        IXLCell FirstCellUsed();

        /// <summary>Returns the first cell with a value of this range.</summary>
        /// <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
        /// <param name="includeStyles">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCell FirstCellUsed(Boolean includeStyles);
        /// <summary>
        /// Returns the last cell of this range.
        /// </summary>
        IXLCell LastCell();
        /// <summary>
        /// Returns the last cell with a value of this range.
        /// <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
        /// </summary>
        IXLCell LastCellUsed();

        /// <summary>Returns the last cell with a value of this range.</summary>
        /// <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
        /// <param name="includeStyles">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCell LastCellUsed(Boolean includeStyles);

        /// <summary>
        /// Determines whether this range contains the specified range address.
        /// </summary>
        /// <param name="rangeAddress">The range address.</param>
        /// <returns>
        ///   <c>true</c> if this range contains the specified range address; otherwise, <c>false</c>.
        /// </returns>
        Boolean ContainsRange(String rangeAddress);
        /// <summary>
        /// Unmerges this range.
        /// </summary>
        IXLRange Unmerge();
        /// <summary>
        /// Merges this range.
        /// <para>The contents and style of the merged cells will be equal to the first cell.</para>
        /// </summary>
        IXLRange Merge();
        /// <summary>
        /// Creates a named range out of this range. 
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <para>The default scope for the named range is Workbook.</para>
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        IXLRange AddToNamed(String rangeName);

        /// <summary>
        /// Creates a named range out of this range. 
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        IXLRange AddToNamed(String rangeName, XLScope scope);

        /// <summary>
        /// Creates a named range out of this range. 
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// <param name="comment">The comments for the named range.</param>
        IXLRange AddToNamed(String rangeName, XLScope scope, String comment);
    }
}
