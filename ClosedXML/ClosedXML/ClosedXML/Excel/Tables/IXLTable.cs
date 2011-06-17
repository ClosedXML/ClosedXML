using System;

namespace ClosedXML.Excel
{
    public enum XLTableTheme
    {
        TableStyleMedium28,
        TableStyleMedium27,
        TableStyleMedium26,
        TableStyleMedium25,
        TableStyleMedium24,
        TableStyleMedium23,
        TableStyleMedium22,
        TableStyleMedium21,
        TableStyleMedium20,
        TableStyleMedium19,
        TableStyleMedium18,
        TableStyleMedium17,
        TableStyleMedium16,
        TableStyleMedium15,
        TableStyleMedium14,
        TableStyleMedium13,
        TableStyleMedium12,
        TableStyleMedium11,
        TableStyleMedium10,
        TableStyleMedium9,
        TableStyleMedium8,
        TableStyleMedium7,
        TableStyleMedium6,
        TableStyleMedium5,
        TableStyleMedium4,
        TableStyleMedium3,
        TableStyleMedium2,
        TableStyleMedium1,
        TableStyleLight21,
        TableStyleLight20,
        TableStyleLight19,
        TableStyleLight18,
        TableStyleLight17,
        TableStyleLight16,
        TableStyleLight15,
        TableStyleLight14,
        TableStyleLight13,
        TableStyleLight12,
        TableStyleLight11,
        TableStyleLight10,
        TableStyleLight9,
        TableStyleLight8,
        TableStyleLight7,
        TableStyleLight6,
        TableStyleLight5,
        TableStyleLight4,
        TableStyleLight3,
        TableStyleLight2,
        TableStyleLight1,
        TableStyleDark11,
        TableStyleDark10,
        TableStyleDark9,
        TableStyleDark8,
        TableStyleDark7,
        TableStyleDark6,
        TableStyleDark5,
        TableStyleDark4,
        TableStyleDark3,
        TableStyleDark2,
        TableStyleDark1
    }
    public interface IXLTable: IXLRangeBase
    {
        String Name { get; set; }
        Boolean EmphasizeFirstColumn { get; set; }
        Boolean EmphasizeLastColumn { get; set; }
        Boolean ShowRowStripes { get; set; }
        Boolean ShowColumnStripes { get; set; }
        Boolean ShowTotalsRow { get; set; }
        Boolean ShowAutoFilter { get; set; }
        XLTableTheme Theme { get; set; }
        IXLRangeRow HeadersRow();
        IXLRangeRow TotalsRow();
        IXLTableField Field(String fieldName);
        IXLTableField Field(Int32 fieldIndex);

        /// <summary>
        /// Gets the first data row of the table.
        /// </summary>
         IXLTableRow FirstRow();
        /// <summary>
        /// Gets the first data row of the table that contains a cell with a value.
        /// </summary>
         IXLTableRow FirstRowUsed();
        /// <summary>
        /// Gets the last data row of the table.
        /// </summary>
         IXLTableRow LastRow();
        /// <summary>
        /// Gets the last data row of the table that contains a cell with a value.
        /// </summary>
         IXLTableRow LastRowUsed();
        /// <summary>
        /// Gets the specified row of the table data.
        /// </summary>
        /// <param name="row">The table row.</param>
         IXLTableRow Row(int row);
        /// <summary>
        /// Gets a collection of all data rows in this table.
        /// </summary>
         IXLTableRows Rows();
        /// <summary>
        /// Gets a collection of the specified data rows in this table.
        /// </summary>
        /// <param name="firstRow">The first row to return.</param>
        /// <param name="lastRow">The last row to return.</param>
        /// <returns></returns>
         IXLTableRows Rows(int firstRow, int lastRow);
        /// <summary>
        /// Gets a collection of the specified data rows in this table, separated by commas.
        /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
        /// </summary>
        /// <param name="rows">The rows to return.</param>
         IXLTableRows Rows(string rows);

         IXLRange Sort();
         IXLRange Sort(Boolean matchCase);
         IXLRange Sort(XLSortOrder sortOrder);
         IXLRange Sort(XLSortOrder sortOrder, Boolean matchCase);
         IXLRange Sort(String columnsToSortBy);
         IXLRange Sort(String columnsToSortBy, Boolean matchCase);

         /// <summary>
         /// Gets the cell at the specified row and column.
         /// <para>The cell address is relative to the parent range.</para>
         /// </summary>
         /// <param name="row">The cell's row.</param>
         /// <param name="column">The cell's column.</param>
         IXLCell Cell(int row, int column);

         /// <summary>Gets the cell at the specified address.</summary>
         /// <para>The cell address is relative to the parent range.</para>
         /// <param name="cellAddressInRange">The cell address in the parent range.</param>
         IXLCell Cell(string cellAddressInRange);

         /// <summary>
         /// Gets the cell at the specified row and column.
         /// <para>The cell address is relative to the parent range.</para>
         /// </summary>
         /// <param name="row">The cell's row.</param>
         /// <param name="column">The cell's column.</param>
         IXLCell Cell(int row, string column);
         /// <summary>Gets the cell at the specified address.</summary>
         /// <para>The cell address is relative to the parent range.</para>
         /// <param name="cellAddressInRange">The cell address in the parent range.</param>
         IXLCell Cell(IXLAddress cellAddressInRange);

         /// <summary>
         /// Gets the specified column of the range.
         /// </summary>
         /// <param name="column">The range column.</param>
         IXLRangeColumn Column(int column);
         /// <summary>
         /// Gets the specified column of the range.
         /// </summary>
         /// <param name="column">The range column.</param>
         IXLRangeColumn Column(string column);
         /// <summary>
         /// Gets the first column of the range.
         /// </summary>
         IXLRangeColumn FirstColumn();
         /// <summary>
         /// Gets the first column of the range that contains a cell with a value.
         /// </summary>
         IXLRangeColumn FirstColumnUsed();
         /// <summary>
         /// Gets the last column of the range.
         /// </summary>
         IXLRangeColumn LastColumn();
         /// <summary>
         /// Gets the last column of the range that contains a cell with a value.
         /// </summary>
         IXLRangeColumn LastColumnUsed();
         /// <summary>
         /// Gets a collection of all columns in this range.
         /// </summary>
         IXLRangeColumns Columns();
         /// <summary>
         /// Gets a collection of the specified columns in this range.
         /// </summary>
         /// <param name="firstColumn">The first column to return.</param>
         /// <param name="lastColumn">The last column to return.</param>
         IXLRangeColumns Columns(int firstColumn, int lastColumn);
         /// <summary>
         /// Gets a collection of the specified columns in this range.
         /// </summary>
         /// <param name="firstColumn">The first column to return.</param>
         /// <param name="lastColumn">The last column to return.</param>
         IXLRangeColumns Columns(string firstColumn, string lastColumn);
         /// <summary>
         /// Gets a collection of the specified columns in this range, separated by commas.
         /// <para>e.g. Columns("G:H"), Columns("10:11,13:14"), Columns("P:Q,S:T"), Columns("V")</para>
         /// </summary>
         /// <param name="columns">The columns to return.</param>
         IXLRangeColumns Columns(string columns);

         /// <summary>
         /// Returns the specified range.
         /// </summary>
         /// <param name="rangeAddress">The range boundaries.</param>
         IXLRange Range(IXLRangeAddress rangeAddress);

         /// <summary>Returns the specified range.</summary>
         /// <para>e.g. Range("A1"), Range("A1:C2")</para>
         /// <param name="rangeAddress">The range boundaries.</param>
         IXLRange Range(string rangeAddress);

         /// <summary>Returns the specified range.</summary>
         /// <param name="firstCell">The first cell in the range.</param>
         /// <param name="lastCell"> The last cell in the range.</param>
         IXLRange Range(IXLCell firstCell, IXLCell lastCell);

         /// <summary>Returns the specified range.</summary>
         /// <param name="firstCellAddress">The first cell address in the range.</param>
         /// <param name="lastCellAddress"> The last cell address in the range.</param>
         IXLRange Range(string firstCellAddress, string lastCellAddress);

         /// <summary>Returns the specified range.</summary>
         /// <param name="firstCellAddress">The first cell address in the range.</param>
         /// <param name="lastCellAddress"> The last cell address in the range.</param>
         IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress);

         /// <summary>Returns a collection of ranges, separated by commas.</summary>
         /// <para>e.g. Ranges("A1"), Ranges("A1:C2"), Ranges("A1:B2,D1:D4")</para>
         /// <param name="ranges">The ranges to return.</param>
         IXLRanges Ranges(string ranges);

         /// <summary>Returns the specified range.</summary>
         /// <param name="firstCellRow">   The first cell's row of the range to return.</param>
         /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
         /// <param name="lastCellRow">    The last cell's row of the range to return.</param>
         /// <param name="lastCellColumn"> The last cell's column of the range to return.</param>
         /// <returns>.</returns>
         IXLRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

         /// <summary>Gets the number of rows in this range.</summary>
         int RowCount();

         /// <summary>Gets the number of columns in this range.</summary>
         int ColumnCount();

         /// <summary>
         /// Inserts X number of columns to the right of this range.
         /// <para>All cells to the right of this range will be shifted X number of columns.</para>
         /// </summary>
         /// <param name="numberOfColumns">Number of columns to insert.</param>
         IXLRangeColumns InsertColumnsAfter(int numberOfColumns);
         IXLRangeColumns InsertColumnsAfter(int numberOfColumns, Boolean expandRange);
         /// <summary>
         /// Inserts X number of columns to the left of this range.
         /// <para>This range and all cells to the right of this range will be shifted X number of columns.</para>
         /// </summary>
         /// <param name="numberOfColumns">Number of columns to insert.</param>
         IXLRangeColumns InsertColumnsBefore(int numberOfColumns);
         IXLRangeColumns InsertColumnsBefore(int numberOfColumns, Boolean expandRange);
         /// <summary>
         /// Inserts X number of rows on top of this range.
         /// <para>This range and all cells below this range will be shifted X number of rows.</para>
         /// </summary>
         /// <param name="numberOfRows">Number of rows to insert.</param>
         IXLRangeRows InsertRowsAbove(int numberOfRows);
         IXLRangeRows InsertRowsAbove(int numberOfRows, Boolean expandRange);
         /// <summary>
         /// Inserts X number of rows below this range.
         /// <para>All cells below this range will be shifted X number of rows.</para>
         /// </summary>
         /// <param name="numberOfRows">Number of rows to insert.</param>
         IXLRangeRows InsertRowsBelow(int numberOfRows);
         IXLRangeRows InsertRowsBelow(int numberOfRows, Boolean expandRange);

         /// <summary>
         /// Deletes this range and shifts the surrounding cells accordingly.
         /// </summary>
         /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
         void Delete(XLShiftDeletedCells shiftDeleteCells);

         /// <summary>
         /// Transposes the contents and styles of all cells in this range.
         /// </summary>
         /// <param name="transposeOption">How to handle the surrounding cells when transposing the range.</param>
         void Transpose(XLTransposeOptions transposeOption);

         IXLTable AsTable();
         IXLTable AsTable(String name);
         IXLTable CreateTable();
         IXLTable CreateTable(String name);

         IXLRange RangeUsed();

         void CopyTo(IXLCell target);
         void CopyTo(IXLRangeBase target);

         void SetAutoFilter();
         void SetAutoFilter(Boolean autoFilter);

         IXLSortElements SortRows { get; }
         IXLSortElements SortColumns { get; }

         
    }
}
