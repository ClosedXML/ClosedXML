// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;

namespace ClosedXML.Excel
{
    public interface IXLWorkbook : IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>, IDisposable
    {
        string Author { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's calculation mode.
        /// </summary>
        XLCalculateMode CalculateMode { get; set; }

        bool CalculationOnSave { get; set; }

        /// <summary>
        ///   Gets or sets the default column width for the workbook.
        ///   <para>All new worksheets will use this column width.</para>
        /// </summary>
        double ColumnWidth { get; set; }

        IXLCustomProperties CustomProperties { get; }

        bool DefaultRightToLeft { get; }

        bool DefaultShowFormulas { get; }

        bool DefaultShowGridLines { get; }

        bool DefaultShowOutlineSymbols { get; }

        bool DefaultShowRowColHeaders { get; }

        bool DefaultShowRuler { get; }

        bool DefaultShowWhiteSpace { get; }

        bool DefaultShowZeros { get; }

        IXLFileSharing FileSharing { get; }

        bool ForceFullCalculation { get; set; }

        bool FullCalculationOnLoad { get; set; }

        bool FullPrecision { get; set; }

        //Boolean IsPasswordProtected { get; }

        //Boolean IsProtected { get; }

        bool LockStructure { get; set; }

        bool LockWindows { get; set; }

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

        /// <summary>
        ///   Gets or sets the workbook's properties.
        /// </summary>
        XLWorkbookProperties Properties { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's reference style.
        /// </summary>
        XLReferenceStyle ReferenceStyle { get; set; }

        bool RightToLeft { get; set; }

        /// <summary>
        ///   Gets or sets the default row height for the workbook.
        ///   <para>All new worksheets will use this row height.</para>
        /// </summary>
        double RowHeight { get; set; }

        bool ShowFormulas { get; set; }

        bool ShowGridLines { get; set; }

        bool ShowOutlineSymbols { get; set; }

        bool ShowRowColHeaders { get; set; }

        bool ShowRuler { get; set; }

        bool ShowWhiteSpace { get; set; }

        bool ShowZeros { get; set; }

        /// <summary>
        ///   Gets or sets the default style for the workbook.
        ///   <para>All new worksheets will use this style.</para>
        /// </summary>
        IXLStyle Style { get; set; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's theme.
        /// </summary>
        IXLTheme Theme { get; }

        bool Use1904DateSystem { get; set; }

        /// <summary>
        ///   Gets an object to manipulate the worksheets.
        /// </summary>
        IXLWorksheets Worksheets { get; }

        IXLWorksheet AddWorksheet();

        IXLWorksheet AddWorksheet(int position);

        IXLWorksheet AddWorksheet(string sheetName);

        IXLWorksheet AddWorksheet(string sheetName, int position);

        IXLWorksheet AddWorksheet(DataTable dataTable);

        void AddWorksheet(DataSet dataSet);

        void AddWorksheet(IXLWorksheet worksheet);

        IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName);

        IXLCell Cell(string namedCell);

        IXLCells Cells(string namedCells);

        IXLCustomProperty CustomProperty(string name);

        object Evaluate(string expression);

        IXLCells FindCells(Func<IXLCell, bool> predicate);

        IXLColumns FindColumns(Func<IXLColumn, bool> predicate);

        IXLRows FindRows(Func<IXLRow, bool> predicate);

        IXLNamedRange NamedRange(string rangeName);

        [Obsolete("Use Protect(String password, Algorithm algorithm, TElement allowedElements)")]
        IXLWorkbookProtection Protect(bool lockStructure, bool lockWindows, string password);

        [Obsolete("Use Protect(String password, Algorithm algorithm, TElement allowedElements)")]
        IXLWorkbookProtection Protect(bool lockStructure);

        [Obsolete("Use Protect(String password, Algorithm algorithm, TElement allowedElements)")]
        IXLWorkbookProtection Protect(bool lockStructure, bool lockWindows);

        IXLRange Range(string range);

        IXLRange RangeFromFullAddress(string rangeAddress, out IXLWorksheet ws);

        IXLRanges Ranges(string ranges);

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
        void Save(bool validate, bool evaluateFormulae = false);

        void Save(SaveOptions options);

        /// <summary>
        ///   Saves the current workbook to a file.
        /// </summary>
        void SaveAs(string file);

        /// <summary>
        ///   Saves the current workbook to a file and optionally validates it.
        /// </summary>
        void SaveAs(string file, bool validate, bool evaluateFormulae = false);

        void SaveAs(string file, SaveOptions options);

        /// <summary>
        ///   Saves the current workbook to a stream.
        /// </summary>
        void SaveAs(Stream stream);

        /// <summary>
        ///   Saves the current workbook to a stream and optionally validates it.
        /// </summary>
        void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false);

        void SaveAs(Stream stream, SaveOptions options);

        /// <summary>
        /// Searches the cells' contents for a given piece of text
        /// </summary>
        /// <param name="searchText">The search text.</param>
        /// <param name="compareOptions">The compare options.</param>
        /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
        /// <returns></returns>
        IEnumerable<IXLCell> Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal, bool searchFormulae = false);

        XLWorkbook SetLockStructure(bool value);

        XLWorkbook SetLockWindows(bool value);

        XLWorkbook SetUse1904DateSystem();

        XLWorkbook SetUse1904DateSystem(bool value);

        /// <summary>
        /// Gets the Excel table of the given name
        /// </summary>
        /// <param name="tableName">Name of the table to return.</param>
        /// <param name="comparisonType">One of the enumeration values that specifies how the strings will be compared.</param>
        /// <returns>The table with given name</returns>
        /// <exception cref="ArgumentOutOfRangeException">If no tables with this name could be found in the workbook.</exception>
        IXLTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase);

        bool TryGetWorksheet(string name, out IXLWorksheet worksheet);

        IXLWorksheet Worksheet(string name);

        IXLWorksheet Worksheet(int position);
    }
}
