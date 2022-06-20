using ClosedXML.Excel.Drawings;
using SkiaSharp;
using System;
using System.IO;

namespace ClosedXML.Excel
{
    public enum XLWorksheetVisibility
    { Visible, Hidden, VeryHidden }

    public interface IXLWorksheet : IXLRangeBase, IXLProtectable<IXLSheetProtection, XLSheetProtectionElements>
    {
        /// <summary>
        /// Gets the workbook that contains this worksheet
        /// </summary>
        XLWorkbook Workbook { get; }

        /// <summary>
        /// Gets or sets the default column width for this worksheet.
        /// </summary>
        double ColumnWidth { get; set; }

        /// <summary>
        /// Gets or sets the default row height for this worksheet.
        /// </summary>
        double RowHeight { get; set; }

        /// <summary>
        /// Gets or sets the name (caption) of this worksheet.
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// Gets or sets the position of the sheet.
        /// <para>When setting the Position all other sheets' positions are shifted accordingly.</para>
        /// </summary>
        int Position { get; set; }

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
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRow FirstRowUsed(bool includeFormats);

        IXLRow FirstRowUsed(XLCellsUsedOptions options);

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
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRow LastRowUsed(bool includeFormats);

        IXLRow LastRowUsed(XLCellsUsedOptions options);

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
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLColumn FirstColumnUsed(bool includeFormats);

        IXLColumn FirstColumnUsed(XLCellsUsedOptions options);

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
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLColumn LastColumnUsed(bool includeFormats);

        IXLColumn LastColumnUsed(XLCellsUsedOptions options);

        /// <summary>
        /// Gets a collection of all columns in this worksheet.
        /// </summary>
        IXLColumns Columns();

        /// <summary>
        /// Gets a collection of the specified columns in this worksheet, separated by commas.
        /// <para>e.g. Columns("G:H"), Columns("10:11,13:14"), Columns("P:Q,S:T"), Columns("V")</para>
        /// </summary>
        /// <param name="columns">The columns to return.</param>
        IXLColumns Columns(string columns);

        /// <summary>
        /// Gets a collection of the specified columns in this worksheet.
        /// </summary>
        /// <param name="firstColumn">The first column to return.</param>
        /// <param name="lastColumn">The last column to return.</param>
        IXLColumns Columns(string firstColumn, string lastColumn);

        /// <summary>
        /// Gets a collection of the specified columns in this worksheet.
        /// </summary>
        /// <param name="firstColumn">The first column to return.</param>
        /// <param name="lastColumn">The last column to return.</param>
        IXLColumns Columns(int firstColumn, int lastColumn);

        /// <summary>
        /// Gets a collection of all rows in this worksheet.
        /// </summary>
        IXLRows Rows();

        /// <summary>
        /// Gets a collection of the specified rows in this worksheet, separated by commas.
        /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
        /// </summary>
        /// <param name="rows">The rows to return.</param>
        IXLRows Rows(string rows);

        /// <summary>
        /// Gets a collection of the specified rows in this worksheet.
        /// </summary>
        /// <param name="firstRow">The first row to return.</param>
        /// <param name="lastRow">The last row to return.</param>
        /// <returns></returns>
        IXLRows Rows(int firstRow, int lastRow);

        /// <summary>
        /// Gets the specified row of the worksheet.
        /// </summary>
        /// <param name="row">The worksheet's row.</param>
        IXLRow Row(int row);

        /// <summary>
        /// Gets the specified column of the worksheet.
        /// </summary>
        /// <param name="column">The worksheet's column.</param>
        IXLColumn Column(int column);

        /// <summary>
        /// Gets the specified column of the worksheet.
        /// </summary>
        /// <param name="column">The worksheet's column.</param>
        IXLColumn Column(string column);

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
        IXLWorksheet CollapseRows(int outlineLevel);

        /// <summary>
        /// Collapses the outlined columns of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet CollapseColumns(int outlineLevel);

        /// <summary>
        /// Expands the outlined rows of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet ExpandRows(int outlineLevel);

        /// <summary>
        /// Expands the outlined columns of the specified level.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLWorksheet ExpandColumns(int outlineLevel);

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
        IXLNamedRange NamedRange(string rangeName);

        /// <summary>
        /// Gets an object to manage how the worksheet is going to displayed by Excel.
        /// </summary>
        IXLSheetView SheetView { get; }

        /// <summary>
        /// Gets the Excel table of the given index
        /// </summary>
        /// <param name="index">Index of the table to return</param>
        IXLTable Table(int index);

        /// <summary>
        /// Gets the Excel table of the given name
        /// </summary>
        /// <param name="name">Name of the table to return</param>
        IXLTable Table(string name);

        /// <summary>
        /// Gets an object to manage this worksheet's Excel tables
        /// </summary>
        IXLTables Tables { get; }

        /// <summary>
        /// Copies the
        /// </summary>
        /// <param name="newSheetName"></param>
        /// <returns></returns>
        IXLWorksheet CopyTo(string newSheetName);

        IXLWorksheet CopyTo(string newSheetName, int position);

        IXLWorksheet CopyTo(XLWorkbook workbook);

        IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName);

        IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName, int position);

        IXLRange RangeUsed();

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRange RangeUsed(bool includeFormats);

        IXLRange RangeUsed(XLCellsUsedOptions options);

        IXLDataValidations DataValidations { get; }

        XLWorksheetVisibility Visibility { get; set; }

        IXLWorksheet Hide();

        IXLWorksheet Unhide();

        IXLSortElements SortRows { get; }

        IXLSortElements SortColumns { get; }

        IXLRange Sort();

        IXLRange Sort(string columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        IXLRange Sort(int columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        //IXLCharts Charts { get; }

        bool ShowFormulas { get; set; }

        bool ShowGridLines { get; set; }

        bool ShowOutlineSymbols { get; set; }

        bool ShowRowColHeaders { get; set; }

        bool ShowRuler { get; set; }

        bool ShowWhiteSpace { get; set; }

        bool ShowZeros { get; set; }

        IXLWorksheet SetShowFormulas(); IXLWorksheet SetShowFormulas(bool value);

        IXLWorksheet SetShowGridLines(); IXLWorksheet SetShowGridLines(bool value);

        IXLWorksheet SetShowOutlineSymbols(); IXLWorksheet SetShowOutlineSymbols(bool value);

        IXLWorksheet SetShowRowColHeaders(); IXLWorksheet SetShowRowColHeaders(bool value);

        IXLWorksheet SetShowRuler(); IXLWorksheet SetShowRuler(bool value);

        IXLWorksheet SetShowWhiteSpace(); IXLWorksheet SetShowWhiteSpace(bool value);

        IXLWorksheet SetShowZeros(); IXLWorksheet SetShowZeros(bool value);

        XLColor TabColor { get; set; }

        IXLWorksheet SetTabColor(XLColor color);

        bool TabSelected { get; set; }

        bool TabActive { get; set; }

        IXLWorksheet SetTabSelected(); IXLWorksheet SetTabSelected(bool value);

        IXLWorksheet SetTabActive(); IXLWorksheet SetTabActive(bool value);

        IXLPivotTable PivotTable(string name);

        IXLPivotTables PivotTables { get; }

        bool RightToLeft { get; set; }

        IXLWorksheet SetRightToLeft(); IXLWorksheet SetRightToLeft(bool value);

        IXLAutoFilter AutoFilter { get; }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRows RowsUsed(bool includeFormats, Func<IXLRow, bool> predicate = null);

        IXLRows RowsUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents, Func<IXLRow, bool> predicate = null);

        IXLRows RowsUsed(Func<IXLRow, bool> predicate);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLColumns ColumnsUsed(bool includeFormats, Func<IXLColumn, bool> predicate = null);

        IXLColumns ColumnsUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents, Func<IXLColumn, bool> predicate = null);

        IXLColumns ColumnsUsed(Func<IXLColumn, bool> predicate);

        IXLRanges MergedRanges { get; }

        IXLConditionalFormats ConditionalFormats { get; }

        IXLSparklineGroups SparklineGroups { get; }

        IXLRanges SelectedRanges { get; }

        IXLCell ActiveCell { get; set; }

        object Evaluate(string expression);

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        void RecalculateAllFormulas();

        string Author { get; set; }

        IXLPictures Pictures { get; }

        IXLPicture Picture(string pictureName);

        IXLPicture AddPicture(Stream stream);

        IXLPicture AddPicture(Stream stream, string name);

        IXLPicture AddPicture(Stream stream, XLPictureFormat format);

        IXLPicture AddPicture(Stream stream, XLPictureFormat format, string name);

        IXLPicture AddPicture(SKCodec bitmap);

        IXLPicture AddPicture(SKCodec bitmap, string name);

        IXLPicture AddPicture(string imageFile);

        IXLPicture AddPicture(string imageFile, string name);
    }
}
