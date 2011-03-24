using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLWorksheetVisibility { Visible, Hidden, VeryHidden }
    public interface IXLWorksheet: IXLRangeBase
    {
        /// <summary>
        /// Gets or sets the default column width for this worksheet.
        /// </summary>
        Double ColumnWidth { get; set; }
        /// <summary>
        /// Gets or sets the default row height for this worksheet.
        /// </summary>
        Double RowHeight { get; set; }

        /// <summary>
        /// Gets or sets the name (caption) of this worksheet.
        /// </summary>
        String Name { get; set; }

        /// <summary>
        /// Gets or sets the position of the sheet.
        /// <para>When setting the Position all other sheets' positions are shifted accordingly.</para>
        /// </summary>
        Int32 Position { get; set; }

        /// <summary>
        /// Gets an object to manipulate the sheet's print options.
        /// </summary>
        IXLPageSetup PageSetup { get; }
        /// <summary>
        /// Gets an object to manipulate the Outline levels.
        /// </summary>
        IXLOutline Outline { get; }

        /// <summary>
        /// Gets the first row of the worksheet.
        /// </summary>
        IXLRow FirstRow();
        /// <summary>
        /// Gets the first row of the worksheet that contains a cell with a value.
        /// </summary>
        IXLRow FirstRowUsed();
        /// <summary>
        /// Gets the last row of the worksheet.
        /// </summary>
        IXLRow LastRow();
        /// <summary>
        /// Gets the last row of the worksheet that contains a cell with a value.
        /// </summary>
        IXLRow LastRowUsed();
        /// <summary>
        /// Gets the first column of the worksheet.
        /// </summary>
        IXLColumn FirstColumn();
        /// <summary>
        /// Gets the first column of the worksheet that contains a cell with a value.
        /// </summary>
        IXLColumn FirstColumnUsed();
        /// <summary>
        /// Gets the last column of the worksheet.
        /// </summary>
        IXLColumn LastColumn();
        /// <summary>
        /// Gets the last column of the worksheet that contains a cell with a value.
        /// </summary>
        IXLColumn LastColumnUsed();
        /// <summary>
        /// Gets a collection of all columns in this worksheet.
        /// </summary>
        IXLColumns Columns();
        /// <summary>
        /// Gets a collection of the specified columns in this worksheet, separated by commas.
        /// <para>e.g. Columns("G:H"), Columns("10:11,13:14"), Columns("P:Q,S:T"), Columns("V")</para>
        /// </summary>
        /// <param name="columns">The columns to return.</param>
        IXLColumns Columns(String columns);
        /// <summary>
        /// Gets a collection of the specified columns in this worksheet.
        /// </summary>
        /// <param name="firstColumn">The first column to return.</param>
        /// <param name="lastColumn">The last column to return.</param>
        IXLColumns Columns(String firstColumn, String lastColumn);
        /// <summary>
        /// Gets a collection of the specified columns in this worksheet.
        /// </summary>
        /// <param name="firstColumn">The first column to return.</param>
        /// <param name="lastColumn">The last column to return.</param>
        IXLColumns Columns(Int32 firstColumn, Int32 lastColumn);
        /// <summary>
        /// Gets a collection of all rows in this worksheet.
        /// </summary>
        IXLRows Rows();
        /// <summary>
        /// Gets a collection of the specified rows in this worksheet, separated by commas.
        /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
        /// </summary>
        /// <param name="rows">The rows to return.</param>
        IXLRows Rows(String rows);
        /// <summary>
        /// Gets a collection of the specified rows in this worksheet.
        /// </summary>
        /// <param name="firstRow">The first row to return.</param>
        /// <param name="lastRow">The last row to return.</param>
        /// <returns></returns>
        IXLRows Rows(Int32 firstRow, Int32 lastRow);
        /// <summary>
        /// Gets the specified row of the worksheet.
        /// </summary>
        /// <param name="row">The worksheet's row.</param>
        IXLRow Row(Int32 row);
        /// <summary>
        /// Gets the specified column of the worksheet.
        /// </summary>
        /// <param name="column">The worksheet's column.</param>
        IXLColumn Column(Int32 column);
        /// <summary>
        /// Gets the specified column of the worksheet.
        /// </summary>
        /// <param name="column">The worksheet's column.</param>
        IXLColumn Column(String column);
        /// <summary>
        /// Gets the cell at the specified row and column.
        /// </summary>
        /// <param name="row">The cell's row.</param>
        /// <param name="column">The cell's column.</param>
        IXLCell Cell(int row, int column);
        /// <summary>Gets the cell at the specified address.</summary>
        /// <param name="cellAddressInRange">The cell address in the worksheet.</param>
        IXLCell Cell(string cellAddressInRange);
        /// <summary>
        /// Gets the cell at the specified row and column.
        /// </summary>
        /// <param name="row">The cell's row.</param>
        /// <param name="column">The cell's column.</param>
        IXLCell Cell(int row, string column);
        /// <summary>Gets the cell at the specified address.</summary>
        /// <param name="cellAddressInRange">The cell address in the worksheet.</param>
        IXLCell Cell(IXLAddress cellAddressInRange);

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
        /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
        /// <param name="lastCellAddress"> The last cell address in the worksheet.</param>
        IXLRange Range(string firstCellAddress, string lastCellAddress);

        /// <summary>Returns the specified range.</summary>
        /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
        /// <param name="lastCellAddress"> The last cell address in the worksheet.</param>
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
        /// Collapses all outlined rows.
        /// </summary>
        IXLWorksheet CollapseRows();
        /// <summary>
        /// Collapses all outlined columns.
        /// </summary>
        IXLWorksheet CollapseColumns();
        /// <summary>
        /// Expands all outlined rows.
        /// </summary>
        IXLWorksheet ExpandRows();
        /// <summary>
        /// Expands all outlined columns.
        /// </summary>
        IXLWorksheet ExpandColumns();

        /// <summary>
        /// Collapses the outlined rows of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet CollapseRows(Int32 outlineLevel);
        /// <summary>
        /// Collapses the outlined columns of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet CollapseColumns(Int32 outlineLevel);
        /// <summary>
        /// Expands the outlined rows of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet ExpandRows(Int32 outlineLevel);
        /// <summary>
        /// Expands the outlined columns of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet ExpandColumns(Int32 outlineLevel);

        /// <summary>
        /// Deletes this worksheet.
        /// </summary>
        void Delete();


        /// <summary>
        /// Gets an object to manage this worksheet's named ranges.
        /// </summary>
        IXLNamedRanges NamedRanges { get; }
        /// <summary>
        /// Gets the specified named range.
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        IXLNamedRange NamedRange(String rangeName);

        /// <summary>
        /// Gets an object to manage how the worksheet is going to displayed by Excel.
        /// </summary>
        IXLSheetView SheetView { get; }

        IXLTable Table(String name);
        IXLTables Tables { get; }

        IXLWorksheet CopyTo(String newSheetName);
        IXLWorksheet CopyTo(String newSheetName, Int32 position);
        IXLWorksheet CopyTo(XLWorkbook workbook, String newSheetName);
        IXLWorksheet CopyTo(XLWorkbook workbook, String newSheetName, Int32 position);

        IXLRange RangeUsed();

        IXLDataValidations DataValidations { get; }

        XLWorksheetVisibility Visibility { get; set; }
        IXLWorksheet Hide();
        IXLWorksheet Unhide();
    }
}
