using ClosedXML.Excel.Drawings;
using System;
using System.Drawing;
using System.IO;

namespace ClosedXML.Excel
{
    public enum XLWorksheetVisibility { Visible, Hidden, VeryHidden }

    public interface IXLWorksheet : IXLRangeBase, IDisposable
    {
        /// <summary>
        /// Gets the workbook that contains this worksheet
        /// </summary>
        XLWorkbook Workbook { get; }

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
        /// <para>Formatted empty cells do not count.</para>
        /// </summary>
        IXLRow FirstRowUsed();

        /// <summary>
        /// Gets the first row of the worksheet that contains a cell with a value.
        /// </summary>
        /// <param name="includeFormats">If set to <c>true</c> formatted empty cells will count as used.</param>
        IXLRow FirstRowUsed(Boolean includeFormats);

        /// <summary>
        /// Gets the last row of the worksheet.
        /// </summary>
        IXLRow LastRow();

        /// <summary>
        /// Gets the last row of the worksheet that contains a cell with a value.
        /// </summary>
        IXLRow LastRowUsed();

        /// <summary>
        /// Gets the last row of the worksheet that contains a cell with a value.
        /// </summary>
        /// <param name="includeFormats">If set to <c>true</c> formatted empty cells will count as used.</param>
        IXLRow LastRowUsed(Boolean includeFormats);

        /// <summary>
        /// Gets the first column of the worksheet.
        /// </summary>
        IXLColumn FirstColumn();

        /// <summary>
        /// Gets the first column of the worksheet that contains a cell with a value.
        /// </summary>
        IXLColumn FirstColumnUsed();

        /// <summary>
        /// Gets the first column of the worksheet that contains a cell with a value.
        /// </summary>
        /// <param name="includeFormats">If set to <c>true</c> formatted empty cells will count as used.</param>
        IXLColumn FirstColumnUsed(Boolean includeFormats);

        /// <summary>
        /// Gets the last column of the worksheet.
        /// </summary>
        IXLColumn LastColumn();

        /// <summary>
        /// Gets the last column of the worksheet that contains a cell with a value.
        /// </summary>
        IXLColumn LastColumnUsed();

        /// <summary>
        /// Gets the last column of the worksheet that contains a cell with a value.
        /// </summary>
        /// <param name="includeFormats">If set to <c>true</c> formatted empty cells will count as used.</param>
        IXLColumn LastColumnUsed(Boolean includeFormats);

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

        /// <summary>Gets the number of rows in this worksheet.</summary>
        int RowCount();

        /// <summary>Gets the number of columns in this worksheet.</summary>
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

        /// <summary>
        /// Gets the Excel table of the given index
        /// </summary>
        /// <param name="index">Index of the table to return</param>
        IXLTable Table(Int32 index);

        /// <summary>
        /// Gets the Excel table of the given name
        /// </summary>
        /// <param name="name">Name of the table to return</param>
        IXLTable Table(String name);

        /// <summary>
        /// Gets an object to manage this worksheet's Excel tables
        /// </summary>
        IXLTables Tables { get; }

        /// <summary>
        /// Copies the
        /// </summary>
        /// <param name="newSheetName"></param>
        /// <returns></returns>
        IXLWorksheet CopyTo(String newSheetName);

        IXLWorksheet CopyTo(String newSheetName, Int32 position);

        IXLWorksheet CopyTo(XLWorkbook workbook, String newSheetName);

        IXLWorksheet CopyTo(XLWorkbook workbook, String newSheetName, Int32 position);

        IXLRange RangeUsed();

        IXLRange RangeUsed(bool includeFormats);

        IXLDataValidations DataValidations { get; }

        XLWorksheetVisibility Visibility { get; set; }

        IXLWorksheet Hide();

        IXLWorksheet Unhide();

        IXLSheetProtection Protection { get; }

        IXLSheetProtection Protect();

        IXLSheetProtection Protect(String password);

        IXLSheetProtection Unprotect();

        IXLSheetProtection Unprotect(String password);

        IXLSortElements SortRows { get; }
        IXLSortElements SortColumns { get; }

        IXLRange Sort();

        IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRange Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        //IXLCharts Charts { get; }

        Boolean ShowFormulas { get; set; }
        Boolean ShowGridLines { get; set; }
        Boolean ShowOutlineSymbols { get; set; }
        Boolean ShowRowColHeaders { get; set; }
        Boolean ShowRuler { get; set; }
        Boolean ShowWhiteSpace { get; set; }
        Boolean ShowZeros { get; set; }

        IXLWorksheet SetShowFormulas(); IXLWorksheet SetShowFormulas(Boolean value);

        IXLWorksheet SetShowGridLines(); IXLWorksheet SetShowGridLines(Boolean value);

        IXLWorksheet SetShowOutlineSymbols(); IXLWorksheet SetShowOutlineSymbols(Boolean value);

        IXLWorksheet SetShowRowColHeaders(); IXLWorksheet SetShowRowColHeaders(Boolean value);

        IXLWorksheet SetShowRuler(); IXLWorksheet SetShowRuler(Boolean value);

        IXLWorksheet SetShowWhiteSpace(); IXLWorksheet SetShowWhiteSpace(Boolean value);

        IXLWorksheet SetShowZeros(); IXLWorksheet SetShowZeros(Boolean value);

        XLColor TabColor { get; set; }

        IXLWorksheet SetTabColor(XLColor color);

        Boolean TabSelected { get; set; }
        Boolean TabActive { get; set; }

        IXLWorksheet SetTabSelected(); IXLWorksheet SetTabSelected(Boolean value);

        IXLWorksheet SetTabActive(); IXLWorksheet SetTabActive(Boolean value);

        IXLPivotTable PivotTable(String name);

        IXLPivotTables PivotTables { get; }

        Boolean RightToLeft { get; set; }

        IXLWorksheet SetRightToLeft(); IXLWorksheet SetRightToLeft(Boolean value);

        IXLAutoFilter AutoFilter { get; }

        IXLRows RowsUsed(Boolean includeFormats = false, Func<IXLRow, Boolean> predicate = null);

        IXLRows RowsUsed(Func<IXLRow, Boolean> predicate);

        IXLColumns ColumnsUsed(Boolean includeFormats = false, Func<IXLColumn, Boolean> predicate = null);

        IXLColumns ColumnsUsed(Func<IXLColumn, Boolean> predicate);

        IXLRanges MergedRanges { get; }
        IXLConditionalFormats ConditionalFormats { get; }

        IXLRanges SelectedRanges { get; }
        IXLCell ActiveCell { get; set; }

        Object Evaluate(String expression);

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        void RecalculateAllFormulas();

        String Author { get; set; }

        IXLPictures Pictures { get; }

        IXLPicture Picture(String pictureName);

        IXLPicture AddPicture(Stream stream);

        IXLPicture AddPicture(Stream stream, String name);

        IXLPicture AddPicture(Stream stream, XLPictureFormat format);

        IXLPicture AddPicture(Stream stream, XLPictureFormat format, String name);

        IXLPicture AddPicture(Bitmap bitmap);

        IXLPicture AddPicture(Bitmap bitmap, String name);

        IXLPicture AddPicture(String imageFile);

        IXLPicture AddPicture(String imageFile, String name);
    }
}
