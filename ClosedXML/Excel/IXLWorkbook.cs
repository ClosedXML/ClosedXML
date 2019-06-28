// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;

namespace ClosedXML.Excel
{
    public interface IXLWorkbook : IDisposable
    {
        String Author { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's calculation mode.
        /// </summary>
        XLCalculateMode CalculateMode { get; set; }

        Boolean CalculationOnSave { get; set; }

        /// <summary>
        ///   Gets or sets the default column width for the workbook.
        ///   <para>All new worksheets will use this column width.</para>
        /// </summary>
        Double ColumnWidth { get; set; }

        IXLCustomProperties CustomProperties { get; }
        Boolean DefaultRightToLeft { get; }
        Boolean DefaultShowFormulas { get; }
        Boolean DefaultShowGridLines { get; }
        Boolean DefaultShowOutlineSymbols { get; }
        Boolean DefaultShowRowColHeaders { get; }
        Boolean DefaultShowRuler { get; }
        Boolean DefaultShowWhiteSpace { get; }
        Boolean DefaultShowZeros { get; }
        IXLFileSharing FileSharing { get; }
        Boolean ForceFullCalculation { get; set; }
        Boolean FullCalculationOnLoad { get; set; }
        Boolean FullPrecision { get; set; }
        Boolean IsPasswordProtected { get; }
        Boolean LockStructure { get; set; }
        Boolean LockWindows { get; set; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's named ranges.
        /// </summary>
        IXLNamedRanges NamedRanges { get; }

        /// <summary>
        ///   Gets or sets the default outline options for the workbook.
        ///   <para>All new worksheets will use these outline options.</para>
        /// </summary>
        IXLOutline Outline { get; set; }

        /// <summary>
        ///   Gets or sets the default page options for the workbook.
        ///   <para>All new worksheets will use these page options.</para>
        /// </summary>
        IXLPageSetup PageOptions { get; set; }

        IXLPivotSources PivotSources { get; }

        /// <summary>
        ///   Gets or sets the workbook's properties.
        /// </summary>
        XLWorkbookProperties Properties { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's reference style.
        /// </summary>
        XLReferenceStyle ReferenceStyle { get; set; }

        Boolean RightToLeft { get; set; }

        /// <summary>
        ///   Gets or sets the default row height for the workbook.
        ///   <para>All new worksheets will use this row height.</para>
        /// </summary>
        Double RowHeight { get; set; }

        Boolean ShowFormulas { get; set; }

        Boolean ShowGridLines { get; set; }

        Boolean ShowOutlineSymbols { get; set; }

        Boolean ShowRowColHeaders { get; set; }

        Boolean ShowRuler { get; set; }

        Boolean ShowWhiteSpace { get; set; }

        Boolean ShowZeros { get; set; }

        /// <summary>
        ///   Gets or sets the default style for the workbook.
        ///   <para>All new worksheets will use this style.</para>
        /// </summary>
        IXLStyle Style { get; set; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's theme.
        /// </summary>
        IXLTheme Theme { get; }

        Boolean Use1904DateSystem { get; set; }

        /// <summary>
        ///   Gets an object to manipulate the worksheets.
        /// </summary>
        IXLWorksheets Worksheets { get; }

        IXLWorksheet AddWorksheet();

        IXLWorksheet AddWorksheet(Int32 position);

        IXLWorksheet AddWorksheet(String sheetName);

        IXLWorksheet AddWorksheet(String sheetName, Int32 position);

        IXLWorksheet AddWorksheet(DataTable dataTable);

        void AddWorksheet(DataSet dataSet);

        void AddWorksheet(IXLWorksheet worksheet);

        IXLWorksheet AddWorksheet(DataTable dataTable, String sheetName);

        IXLCell Cell(String namedCell);

        IXLCells Cells(String namedCells);

        IXLCustomProperty CustomProperty(String name);

        Object Evaluate(String expression);

        IXLCells FindCells(Func<IXLCell, Boolean> predicate);

        IXLColumns FindColumns(Func<IXLColumn, Boolean> predicate);

        IXLRows FindRows(Func<IXLRow, Boolean> predicate);

        IXLNamedRange NamedRange(String rangeName);

        void Protect(Boolean lockStructure, Boolean lockWindows, String workbookPassword);

        void Protect();

        void Protect(string workbookPassword);

        void Protect(Boolean lockStructure);

        void Protect(Boolean lockStructure, Boolean lockWindows);

        IXLRange Range(String range);

        IXLRange RangeFromFullAddress(String rangeAddress, out IXLWorksheet ws);

        IXLRanges Ranges(String ranges);

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        void RecalculateAllFormulas();

        /// <summary>
        ///   Saves the current workbook.
        /// </summary>
        void Save();

        /// <summary>
        ///   Saves the current workbook and optionally performs validation
        /// </summary>
        void Save(Boolean validate, Boolean evaluateFormulae = false);

        void Save(SaveOptions options);

        /// <summary>
        ///   Saves the current workbook to a file.
        /// </summary>
        void SaveAs(String file);

        /// <summary>
        ///   Saves the current workbook to a file and optionally validates it.
        /// </summary>
        void SaveAs(String file, Boolean validate, Boolean evaluateFormulae = false);

        void SaveAs(String file, SaveOptions options);

        /// <summary>
        ///   Saves the current workbook to a stream.
        /// </summary>
        void SaveAs(Stream stream);

        /// <summary>
        ///   Saves the current workbook to a stream and optionally validates it.
        /// </summary>
        void SaveAs(Stream stream, Boolean validate, Boolean evaluateFormulae = false);

        void SaveAs(Stream stream, SaveOptions options);

        /// <summary>
        /// Searches the cells' contents for a given piece of text
        /// </summary>
        /// <param name="searchText">The search text.</param>
        /// <param name="compareOptions">The compare options.</param>
        /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
        /// <returns></returns>
        IEnumerable<IXLCell> Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false);

        XLWorkbook SetLockStructure(Boolean value);

        XLWorkbook SetLockWindows(Boolean value);

        XLWorkbook SetUse1904DateSystem();

        XLWorkbook SetUse1904DateSystem(Boolean value);

        /// <summary>
        /// Gets the Excel table of the given name
        /// </summary>
        /// <param name="tableName">Name of the table to return.</param>
        /// <param name="comparisonType">One of the enumeration values that specifies how the strings will be compared.</param>
        /// <returns>The table with given name</returns>
        /// <exception cref="ArgumentOutOfRangeException">If no tables with this name could be found in the workbook.</exception>
        IXLTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase);

        Boolean TryGetWorksheet(String name, out IXLWorksheet worksheet);

        void Unprotect();

        void Unprotect(string workbookPassword);

        IXLWorksheet Worksheet(String name);

        IXLWorksheet Worksheet(Int32 position);
    }
}
