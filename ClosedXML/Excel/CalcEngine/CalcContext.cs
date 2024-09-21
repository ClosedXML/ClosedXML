using ClosedXML.Excel.CalcEngine.Exceptions;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcContext
    {
        private readonly XLCalcEngine _calcEngine;
        private readonly XLWorkbook? _workbook;
        private readonly XLWorksheet? _worksheet;
        private readonly IXLAddress? _formulaAddress;
        private readonly bool _recursive;

        public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLCell cell)
            : this(calcEngine, culture, cell.Worksheet.Workbook, cell.Worksheet, cell.Address)
        {
        }

        public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLWorkbook? workbook, XLWorksheet? worksheet, IXLAddress? formulaAddress, bool recursive = false)
        {
            _calcEngine = calcEngine;
            _workbook = workbook;
            _worksheet = worksheet;
            _formulaAddress = formulaAddress;
            _recursive = recursive;
            Culture = culture;
        }

        // LEGACY: Remove once legacy functions are migrated
        internal XLCalcEngine CalcEngine => _calcEngine ?? throw new MissingContextException();

        /// <summary>
        /// Worksheet of the cell the formula is calculating.
        /// </summary>
        public XLWorkbook Workbook => _workbook ?? throw new MissingContextException();

        /// <summary>
        /// Worksheet of the cell the formula is calculating.
        /// </summary>
        public XLWorksheet Worksheet => _worksheet ?? throw new MissingContextException();

        /// <summary>
        /// Address of the calculated formula.
        /// </summary>
        public IXLAddress FormulaAddress => _formulaAddress ?? throw new MissingContextException();

        /// <summary>
        /// A culture used for comparisons and conversions (e.g. text to number).
        /// </summary>
        public CultureInfo Culture { get; }

        /// <summary>
        /// Excel 2016 and earlier doesn't support dynamic array formulas (it used an array formulas instead). As a consequence,
        /// all arguments for scalar functions where passed through implicit intersection before calling the function.
        /// </summary>
        public bool UseImplicitIntersection => true;

        /// <summary>
        /// Should functions be calculated per item of multi-values argument in the scalar parameters.
        /// </summary>
        public bool IsArrayCalculation { get; set; }

        /// <summary>
        /// Sheet that is being recalculated. If set, formula can read dirty
        /// values from other sheets, but not from this sheetId.
        /// </summary>
        public uint? RecalculateSheetId { get; set; }

        internal XLSheetPoint FormulaSheetPoint => new(FormulaAddress.RowNumber, FormulaAddress.ColumnNumber);

        internal ScalarValue GetCellValue(XLWorksheet? sheet, int rowNumber, int columnNumber)
        {
            sheet ??= Worksheet;
            var valueSlice = sheet.Internals.CellsCollection.ValueSlice;
            var point = new XLSheetPoint(rowNumber, columnNumber);
            var formula = sheet.Internals.CellsCollection.FormulaSlice.Get(point);

            if (formula is null)
                return valueSlice.GetCellValue(point);

            if (!formula.IsDirty)
                return valueSlice.GetCellValue(point);

            // Used when only one sheet should be recalculated, leaving other sheets with their data.
            if (RecalculateSheetId is not null && sheet.SheetId != RecalculateSheetId.Value)
                return valueSlice.GetCellValue(point);

            // A special branch for functions out of cells (e.g. worksheet.Evaluate("A1+2")).
            // These functions are not a part of calculation chain and thus reordering a chain
            // for them doesn't make sense.
            if (_recursive)
            {
                var cell = sheet.GetCell(point);
                return cell?.Value ?? Blank.Value;
            }

            throw new GettingDataException(new XLBookPoint(sheet.SheetId, new XLSheetPoint(rowNumber, columnNumber)));
        }

        /// <summary>
        /// Get cells with a value for a reference.
        /// </summary>
        /// <param name="reference">Reference for which to return cells.</param>
        /// <returns>A lazy (non-materialized) enumerable of cells with a value for the reference.</returns>
        internal IEnumerable<XLCell> GetNonBlankCells(Reference reference)
        {
            // XLCells is not suitable here, e.g. it doesn't count a cell twice if it is in multiple areas
            var nonBlankCells = Enumerable.Empty<XLCell>();
            foreach (var area in reference.Areas)
            {
                var areaCells = Worksheet.Internals.CellsCollection
                    .GetCells(
                        area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber,
                        area.LastAddress.RowNumber, area.LastAddress.ColumnNumber,
                        cell => !cell.IsEmpty());
                nonBlankCells = nonBlankCells.Concat(areaCells);
            }

            return nonBlankCells;
        }
    }
}
