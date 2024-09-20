#nullable disable

using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    public enum XLScope
    {
        Workbook,
        Worksheet
    }

    public interface IXLRangeBase : IXLAddressable
    {
        IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Sets a value to every cell in this range.
        /// <para>
        /// Setter will clear a formula, if the cell contains a formula.
        /// If the value is a text that starts with a single quote, setter will prefix the value with a single quote through
        /// <see cref="IXLStyle.IncludeQuotePrefix"/> in Excel too and the value of cell is set to to non-quoted text.
        /// </para>
        /// </summary>
        XLCellValue Value { set; }

        /// <summary>
        ///   Sets the cells' formula with A1 references.
        /// </summary>
        /// <value>The formula with A1 references.</value>
        String FormulaA1 { set; }

        /// <summary>
        /// Create an array formula for all cells in the range.
        /// </summary>
        /// <exception cref="InvalidOperationException">When the range overlaps with a table, pivot table, merged cells or partially overlaps another array formula.</exception>
        String FormulaArrayA1 { set; }

        /// <summary>
        ///   Sets the cells' formula with R1C1 references.
        /// </summary>
        /// <value>The formula with R1C1 references.</value>
        String FormulaR1C1 { set; }

        IXLStyle Style { get; set; }

        /// <summary>
        ///   Gets or sets a value indicating whether this cell's text should be shared or not.
        /// </summary>
        /// <value>
        ///   If false the cell's text will not be shared and stored as an inline value.
        /// </value>
        Boolean ShareString { set; }

        /// <summary>
        ///   Returns the collection of cells.
        /// </summary>
        IXLCells Cells();

        IXLCells Cells(Boolean usedCellsOnly);

        IXLCells Cells(Boolean usedCellsOnly, XLCellsUsedOptions options);

        IXLCells Cells(String cells);

        IXLCells Cells(Func<IXLCell, Boolean> predicate);

        /// <summary>
        ///   Returns the collection of cells that have a value. Formats are ignored.
        /// </summary>
        IXLCells CellsUsed();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        IXLCells CellsUsed(XLCellsUsedOptions options);

        IXLCells CellsUsed(Func<IXLCell, Boolean> predicate);

        IXLCells CellsUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate);

        /// <summary>
        /// Searches the cells' contents for a given piece of text
        /// </summary>
        /// <param name="searchText">The search text.</param>
        /// <param name="compareOptions">The compare options.</param>
        /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
        IXLCells Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false);

        /// <summary>
        ///   Returns the first cell of this range.
        /// </summary>
        IXLCell FirstCell();

        /// <summary>
        ///   Returns the first non-empty cell with a value of this range. Formats are ignored.
        ///   <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
        /// </summary>
        IXLCell FirstCellUsed();

        /// <summary>
        /// Returns the first non-empty cell with a value of this range.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        IXLCell FirstCellUsed(XLCellsUsedOptions options);

        IXLCell FirstCellUsed(Func<IXLCell, Boolean> predicate);

        /// <summary>
        /// Returns the first non-empty cell with a value of this range.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        /// <param name="predicate">The predicate used to choose cells</param>
        IXLCell FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate);

        /// <summary>
        ///   Returns the last cell of this range.
        /// </summary>
        IXLCell LastCell();

        /// <summary>
        ///   Returns the last non-empty cell with a value of this range. Formats are ignored.
        ///   <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
        /// </summary>
        IXLCell LastCellUsed();

        /// <summary>
        /// Returns the last non-empty cell with a value of this range.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        IXLCell LastCellUsed(XLCellsUsedOptions options);

        IXLCell LastCellUsed(Func<IXLCell, Boolean> predicate);

        IXLCell LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate);

        /// <summary>
        ///   Determines whether this range contains the specified range (completely).
        ///   <para>For partial matches use the range.Intersects method.</para>
        /// </summary>
        /// <param name = "rangeAddress">The range address.</param>
        /// <returns>
        ///   <c>true</c> if this range contains the specified range; otherwise, <c>false</c>.
        /// </returns>
        Boolean Contains(String rangeAddress);

        /// <summary>
        ///   Determines whether this range contains the specified range (completely).
        ///   <para>For partial matches use the range.Intersects method.</para>
        /// </summary>
        /// <param name = "range">The range to match.</param>
        /// <returns>
        ///   <c>true</c> if this range contains the specified range; otherwise, <c>false</c>.
        /// </returns>
        Boolean Contains(IXLRangeBase range);

        Boolean Contains(IXLCell cell);

        /// <summary>
        ///   Determines whether this range intersects the specified range.
        ///   <para>For whole matches use the range.Contains method.</para>
        /// </summary>
        /// <param name = "rangeAddress">The range address.</param>
        /// <returns>
        ///   <c>true</c> if this range intersects the specified range; otherwise, <c>false</c>.
        /// </returns>
        Boolean Intersects(String rangeAddress);

        /// <summary>
        ///   Determines whether this range contains the specified range.
        ///   <para>For whole matches use the range.Contains method.</para>
        /// </summary>
        /// <param name = "range">The range to match.</param>
        /// <returns>
        ///   <c>true</c> if this range intersects the specified range; otherwise, <c>false</c>.
        /// </returns>
        Boolean Intersects(IXLRangeBase range);

        /// <summary>
        ///   Unmerges this range.
        /// </summary>
        IXLRange Unmerge();

        /// <summary>
        /// Merges this range. Only the top-left cell will have a value, other values will be blank.
        /// </summary>
        IXLRange Merge();

        IXLRange Merge(Boolean checkIntersect);

        /// <summary>
        /// Creates/adds this range to workbook scoped <see cref="IXLDefinedNames"/>.
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// </summary>
        /// <param name = "name">Name of the defined name, without sheet.</param>
        IXLRange AddToNamed(String name);

        /// <summary>
        /// Creates/adds this range to <see cref="IXLDefinedNames"/>.
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name = "name">Name of the defined name, without sheet.</param>
        /// <param name = "scope">The scope for the named range.</param>
        /// </summary>
        IXLRange AddToNamed(String name, XLScope scope);

        /// <summary>
        /// Creates/adds this range to <see cref="IXLDefinedNames"/>.
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name = "name">Name of the defined name, without sheet.</param>
        /// <param name = "scope">The scope for the named range.</param>
        /// <param name = "comment">The comments for the named range.</param>
        /// </summary>
        IXLRange AddToNamed(String name, XLScope scope, String comment);

        /// <summary>
        /// Clears the contents of this range.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.All);

        /// <summary>
        ///   Deletes the cell comments from this range.
        /// </summary>
        void DeleteComments();

        /// <summary>
        /// Set value to all cells in the range.
        /// </summary>
        IXLRangeBase SetValue(XLCellValue value);

        /// <summary>
        ///   Converts this object to a range.
        /// </summary>
        IXLRange AsRange();

        Boolean IsMerged();

        Boolean IsEmpty();

        Boolean IsEmpty(XLCellsUsedOptions options);

        /// <summary>
        /// Determines whether range address spans the entire column.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire column; otherwise, <c>false</c>.
        /// </returns>
        Boolean IsEntireColumn();

        /// <summary>
        /// Determines whether range address spans the entire row.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire row; otherwise, <c>false</c>.
        /// </returns>

        Boolean IsEntireRow();

        /// <summary>
        /// Determines whether the range address spans the entire worksheet.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire sheet; otherwise, <c>false</c>.
        /// </returns>
        Boolean IsEntireSheet();

        IXLPivotTable CreatePivotTable(IXLCell targetCell, String name);

        //IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn);

        IXLAutoFilter SetAutoFilter();

        IXLAutoFilter SetAutoFilter(Boolean value);

        /// <summary>
        /// Returns a data validation rule assigned to the range, if any, or creates a new instance of data validation rule if no rule exists.
        /// </summary>
        IXLDataValidation GetDataValidation();

        /// <summary>
        /// Creates a new data validation rule for the range, replacing the existing one.
        /// </summary>
        IXLDataValidation CreateDataValidation();

        [Obsolete("Use GetDataValidation() to access the existing rule, or CreateDataValidation() to create a new one.")]
        IXLDataValidation SetDataValidation();

        IXLConditionalFormat AddConditionalFormat();

        void Select();

        /// <summary>
        /// Grows this the current range by one cell to each side
        /// </summary>
        IXLRangeBase Grow();

        /// <summary>
        /// Grows this the current range by the specified number of cells to each side.
        /// </summary>
        /// <param name="growCount">The grow count.</param>
        IXLRangeBase Grow(Int32 growCount);

        /// <summary>
        /// Shrinks this current range by one cell.
        /// </summary>
        IXLRangeBase Shrink();

        /// <summary>
        /// Shrinks the current range by the specified number of cells from each side.
        /// </summary>
        /// <param name="shrinkCount">The shrink count.</param>
        IXLRangeBase Shrink(Int32 shrinkCount);

        /// <summary>
        /// Returns the intersection of this range with another range on the same worksheet.
        /// </summary>
        /// <param name="otherRange">The other range.</param>
        /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
        /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
        /// <returns>The range address of the intersection</returns>
        IXLRangeAddress Intersection(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null);

        /// <summary>
        /// Returns the set of cells surrounding the current range.
        /// </summary>
        /// <param name="predicate">The predicate to apply on the resulting set of cells.</param>
        IXLCells SurroundingCells(Func<IXLCell, Boolean> predicate = null);

        /// <summary>
        /// Calculates the union of two ranges on the same worksheet.
        /// </summary>
        /// <param name="otherRange">The other range.</param>
        /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
        /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
        /// <returns>
        /// The union
        /// </returns>
        IXLCells Union(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null);

        /// <summary>
        /// Returns all cells in the current range that are not in the other range.
        /// </summary>
        /// <param name="otherRange">The other range.</param>
        /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
        /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
        IXLCells Difference(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null);

        /// <summary>
        /// Returns a range so that its offset from the target base range is equal to the offset of the current range to the source base range.
        /// For example, if the current range is D4:E4, the source base range is A1:C3, then the relative range to the target base range B10:D13 is E14:F14
        /// </summary>
        /// <param name="sourceBaseRange">The source base range.</param>
        /// <param name="targetBaseRange">The target base range.</param>
        /// <returns>The relative range</returns>
        IXLRangeBase Relative(IXLRangeBase sourceBaseRange, IXLRangeBase targetBaseRange);
    }
}
