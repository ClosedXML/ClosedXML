using ClosedXML.Excel.CalcEngine.Exceptions;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Parser;
using System;

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
        /// This method goes over slices and returns a value for each non-blank cell. Because it is using
        /// slice iterators, it scales with number of cells, not a size of area in reference (i.e. it works
        /// fine even if reference is <c>A1:XFD1048576</c>). It also works for 3D references.  
        /// </summary>
        internal IEnumerable<ScalarValue> GetNonBlankValues(Reference reference)
        {
            foreach (var area in reference.Areas)
            {
                var sheet = area.Worksheet ?? Worksheet;
                var range = XLSheetRange.FromRangeAddress(area);

                // A value can be either in a non-empty value slice or a empty cell with a formula.
                var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
                while (enumerator.MoveNext())
                {
                    var point = enumerator.Current;
                    var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                    if (!scalarValue.IsBlank)
                        yield return scalarValue;
                }
            }
        }

        internal IEnumerable<ScalarValue> GetFilteredNonBlankValues(Reference reference, string function, bool skipHiddenRows = false)
        {
            // Allocate one per call, because visitor holds info whether function was found in a formula.
            var visitor = new FunctionVisitor(function);
            foreach (var area in reference.Areas)
            {
                var sheet = area.Worksheet ?? Worksheet;
                var range = XLSheetRange.FromRangeAddress(area);
                var currentRow = 0;
                var rowIsHidden = true;

                // A value can be either in a non-empty value slice or a empty cell with a formula.
                var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
                while (enumerator.MoveNext())
                {
                    var point = enumerator.Current;

                    if (skipHiddenRows)
                    {
                        // If row changed, update hidden info about current row
                        if (currentRow != point.Row)
                        {
                            currentRow = point.Row;
                            rowIsHidden = sheet.Internals.RowsCollection.TryGetValue(currentRow, out var row) && row.IsHidden;
                        }

                        if (rowIsHidden)
                            continue;
                    }

                    var formula = sheet.Internals.CellsCollection.FormulaSlice.Get(point);
                    if (CallsFunction(formula, visitor))
                        continue;

                    var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                    if (!scalarValue.IsBlank)
                        yield return scalarValue;
                }
            }

            yield break;

            static bool CallsFunction(XLCellFormula? formula, FunctionVisitor visitor)
            {
                if (formula is null)
                    return false;

                if (!formula.A1.Contains(visitor.FunctionName, StringComparison.OrdinalIgnoreCase))
                    return false;

                FormulaParser<object?, object?, FunctionVisitor>.CellFormulaA1(formula.A1, visitor, visitor);
                if (!visitor.Found)
                    return false;

                // In order to reuse same visitor without allocation, clear the found flag.
                visitor.Clear();
                return true;
            }
        }

        /// <summary>
        /// This method should be used mostly for range arguments. If a value is scalar,
        /// return a single value enumerable.
        /// </summary>
        internal IEnumerable<ScalarValue> GetNonBlankValues(AnyValue value)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
            {
                if (scalar.IsBlank)
                    return System.Array.Empty<ScalarValue>();

                return new ScalarArray(scalar, 1, 1);
            }

            if (collection.TryPickT0(out var array, out var reference))
                return array.Where(x => !x.IsBlank);

            return GetNonBlankValues(reference);
        }

        private class FunctionVisitor : CollectVisitor<FunctionVisitor>
        {
            public FunctionVisitor(string function)
            {
                FunctionName = function;
            }

            internal string FunctionName { get; }

            public bool Found { get; private set; }

            public void Clear() => Found = false;

            public override object? Function(FunctionVisitor context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<object?> arguments)
            {
                Found = Found || functionName.Equals(FunctionName.AsSpan(), StringComparison.OrdinalIgnoreCase);
                return default;
            }
        }
    }
}
