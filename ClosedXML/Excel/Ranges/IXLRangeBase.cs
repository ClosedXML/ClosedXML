using System;

namespace ClosedXML.Excel
{
    public enum XLScope
    {
        Workbook,
        Worksheet
    }

    public interface IXLRangeBase: IDisposable
    {
        IXLWorksheet Worksheet { get; }

        /// <summary>
        ///   Gets an object with the boundaries of this range.
        /// </summary>
        IXLRangeAddress RangeAddress { get; }

        /// <summary>
        ///   Sets a value to every cell in this range.
        ///   <para>If the object is an IEnumerable ClosedXML will copy the collection's data into a table starting from each cell.</para>
        ///   <para>If the object is a range ClosedXML will copy the range starting from each cell.</para>
        ///   <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
        ///   <para>ClosedXML will try to translate it to the corresponding type, if it can't then the value will be left as a string.</para>
        /// </summary>
        /// <value>
        ///   The object containing the value(s) to set.
        /// </value>
        Object Value { set; }

        /// <summary>
        ///   Sets the type of the cells' data.
        ///   <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
        ///   <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
        /// </summary>
        /// <value>
        ///   The type of the cell's data.
        /// </value>
        /// <exception cref = "ArgumentException"></exception>
        XLCellValues DataType { set; }

        /// <summary>
        ///   Sets the cells' formula with A1 references.
        /// </summary>
        /// <value>The formula with A1 references.</value>
        String FormulaA1 { set; }

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

        IXLHyperlinks Hyperlinks { get; }


        /// <summary>
        ///   Returns the collection of cells.
        /// </summary>
        IXLCells Cells();

        IXLCells Cells(Boolean usedCellsOnly);

        IXLCells Cells(Boolean usedCellsOnly, Boolean includeFormats);

        IXLCells Cells(String cells);

        IXLCells Cells(Func<IXLCell, Boolean> predicate);

        /// <summary>
        ///   Returns the collection of cells that have a value. Formats are ignored.
        /// </summary>
        IXLCells CellsUsed();

        /// <summary>
        ///   Returns the collection of cells that have a value.
        /// </summary>
        /// <param name = "includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCells CellsUsed(Boolean includeFormats);

        IXLCells CellsUsed(Func<IXLCell, Boolean> predicate);

        IXLCells CellsUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate);

        /// <summary>
        ///   Returns the first cell of this range.
        /// </summary>
        IXLCell FirstCell();

        /// <summary>
        ///   Returns the first cell with a value of this range. Formats are ignored.
        ///   <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
        /// </summary>
        IXLCell FirstCellUsed();

        /// <summary>
        ///   Returns the first cell with a value of this range.
        /// </summary>
        /// <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
        /// <param name = "includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCell FirstCellUsed(Boolean includeFormats);

        IXLCell FirstCellUsed(Func<IXLCell, Boolean> predicate);

        IXLCell FirstCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate);

        /// <summary>
        ///   Returns the last cell of this range.
        /// </summary>
        IXLCell LastCell();

        /// <summary>
        ///   Returns the last cell with a value of this range. Formats are ignored.
        ///   <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
        /// </summary>
        IXLCell LastCellUsed();

        /// <summary>
        ///   Returns the last cell with a value of this range.
        /// </summary>
        /// <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
        /// <param name = "includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCell LastCellUsed(Boolean includeFormats);

        IXLCell LastCellUsed(Func<IXLCell, Boolean> predicate);

        IXLCell LastCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate);

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
        ///   Merges this range.
        ///   <para>The contents and style of the merged cells will be equal to the first cell.</para>
        /// </summary>
        IXLRange Merge();

        IXLRange Merge(Boolean checkIntersect);

        /// <summary>
        ///   Creates a named range out of this range.
        ///   <para>If the named range exists, it will add this range to that named range.</para>
        ///   <para>The default scope for the named range is Workbook.</para>
        /// </summary>
        /// <param name = "rangeName">Name of the range.</param>
        IXLRange AddToNamed(String rangeName);

        /// <summary>
        ///   Creates a named range out of this range.
        ///   <para>If the named range exists, it will add this range to that named range.</para>
        ///   <param name = "rangeName">Name of the range.</param>
        ///   <param name = "scope">The scope for the named range.</param>
        /// </summary>
        IXLRange AddToNamed(String rangeName, XLScope scope);

        /// <summary>
        ///   Creates a named range out of this range.
        ///   <para>If the named range exists, it will add this range to that named range.</para>
        ///   <param name = "rangeName">Name of the range.</param>
        ///   <param name = "scope">The scope for the named range.</param>
        ///   <param name = "comment">The comments for the named range.</param>
        /// </summary>
        IXLRange AddToNamed(String rangeName, XLScope scope, String comment);

        /// <summary>
        /// Clears the contents of this range.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);

        /// <summary>
        ///   Deletes the cell comments from this range.
        /// </summary>
        void DeleteComments();

        IXLRangeBase SetValue<T>(T value);

        /// <summary>
        ///   Converts this object to a range.
        /// </summary>
        IXLRange AsRange();

        Boolean IsMerged();
        Boolean IsEmpty();
        Boolean IsEmpty(Boolean includeFormats);


        IXLPivotTable CreatePivotTable(IXLCell targetCell);
        IXLPivotTable CreatePivotTable(IXLCell targetCell, String name);



        //IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn);

        IXLAutoFilter SetAutoFilter();

        IXLDataValidation SetDataValidation();

        IXLConditionalFormat AddConditionalFormat();

        void Select();
    }
}
