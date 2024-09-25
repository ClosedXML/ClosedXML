using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Exceptions;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// CalcEngine parses strings and returns Expression objects that can
    /// be evaluated.
    /// </summary>
    /// <remarks>
    /// <para>This class has three extensibility points:</para>
    /// <para>Use the <b>RegisterFunction</b> method to define custom functions.</para>
    /// </remarks>
    internal class XLCalcEngine : ISheetListener, IWorkbookListener
    {
        private readonly CultureInfo _culture;
        private readonly ExpressionCache _cache;               // cache with parsed expressions
        private readonly FormulaParser _parser;
        private readonly CalculationVisitor _visitor;
        private DependencyTree? _dependencyTree;
        private XLCalculationChain? _chain;

        public XLCalcEngine(CultureInfo culture)
        {
            _culture = culture;
            _cache = new ExpressionCache(this);
            var funcRegistry = GetFunctionTable();
            _parser = new FormulaParser(funcRegistry);
            _visitor = new CalculationVisitor(funcRegistry);
            _dependencyTree = null;
            _chain = null;
        }

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public Formula Parse(string expression)
        {
            return _parser.GetAst(expression, isA1: true);
        }

        public Formula ParseR1C1(string expression)
        {
            return _parser.GetAst(expression, isA1: false);
        }

        /// <summary>
        /// Add an array formula to the calc engine to manage dirty tracking and evaluation.
        /// </summary>
        internal void AddArrayFormula(XLSheetRange range, XLCellFormula arrayFormula, XLWorksheet sheet)
        {
            if (_chain is not null && _dependencyTree is not null)
            {
                _dependencyTree.AddFormula(new XLBookArea(sheet.Name, range), arrayFormula, sheet.Workbook);
                _chain.AppendArea(sheet.SheetId, range);
            }
        }

        /// <summary>
        /// Add a formula to the calc engine to manage dirty tracking and evaluation.
        /// </summary>
        internal void AddNormalFormula(XLBookPoint point, string sheetName, XLCellFormula formula, XLWorkbook workbook)
        {
            if (_chain is not null && _dependencyTree is not null)
            {
                var pointArea = new XLBookArea(sheetName, new XLSheetRange(point.Point, point.Point));
                _dependencyTree.AddFormula(pointArea, formula, workbook);
                _chain.AddLast(point);
            }
        }

        /// <summary>
        /// Remove formula from dependency tree (=precedents won't mark
        /// it as dirty) and remove <paramref name="point"/> from the chain.
        /// Note that even if formula is used by many cells (e.g. array formula),
        /// it is fully removed from dependency tree, but each cells referencing
        /// the formula must be removed individually from calc chain.
        /// </summary>
        internal void RemoveFormula(XLBookPoint point, XLCellFormula formula)
        {
            if (_chain is not null && _dependencyTree is not null)
            {
                _dependencyTree.RemoveFormula(formula);
                _chain.Remove(point);
            }
        }

        internal void OnAddedSheet(XLWorksheet sheet)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        internal void OnDeletingSheet(XLWorksheet sheet)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnInsertAreaAndShiftDown(XLWorksheet sheet, XLSheetRange area)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnInsertAreaAndShiftRight(XLWorksheet sheet, XLSheetRange area)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnDeleteAreaAndShiftLeft(XLWorksheet sheet, XLSheetRange deletedArea)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnDeleteAreaAndShiftUp(XLWorksheet sheet, XLSheetRange deletedArea)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        private void Purge(XLWorksheets sheets)
        {
            _dependencyTree = null;
            _chain = null;

            // Mark everything as dirty, because there can be stale values
            foreach (var sheet in sheets)
            {
                sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(XLSheetRange.Full);
            }
        }

        internal void MarkDirty(XLWorksheet sheet, XLSheetPoint point)
        {
            MarkDirty(sheet, new XLSheetRange(point, point));
        }

        internal void MarkDirty(XLWorksheet sheet, XLSheetRange area)
        {
            if (_dependencyTree is not null)
            {
                var bookArea = new XLBookArea(sheet.Name, area);
                _dependencyTree.MarkDirty(bookArea);
            }
        }

        /// <summary>
        /// Recalculate a workbook or a sheet.
        /// </summary>
        internal void Recalculate(XLWorkbook wb, uint? recalculateSheetId)
        {
            // Lazy, so initialize chain from wb, if it is empty
            if (_chain is null || _dependencyTree is null)
            {
                _chain = XLCalculationChain.CreateFrom(wb);
                _dependencyTree = DependencyTree.CreateFrom(wb);
            }

            var sheetIdMap = wb.WorksheetsInternal
                .ToDictionary<XLWorksheet, uint, (XLWorksheet Sheet, ValueSlice ValueSlice, FormulaSlice FormulaSlice)>(
                    sheet => sheet.SheetId,
                    sheet => (sheet, sheet.Internals.CellsCollection.ValueSlice, sheet.Internals.CellsCollection.FormulaSlice));

            // Each outer loop moves chain one cell ahead.
            while (_chain.MoveAhead())
            {
                // Inner loop that pushes supporting formulas ahead of current.
                // It ends when a cell has been calculated and thus chain can move ahead.
                while (true)
                {
                    var current = _chain.Current;
                    var sheetId = current.SheetId;

                    // Skip dirty cells from sheets that are not being recalculated
                    if (recalculateSheetId is not null && sheetId != recalculateSheetId.Value)
                    {
                        // Even though cell is dirty, it's in the ignored sheet and
                        // thus chain can move ahead.
                        break;
                    }

                    if (!sheetIdMap.TryGetValue(sheetId, out var sheetInfo))
                    {
                        throw new InvalidOperationException($"Unable to find sheet with sheetId {sheetId} for a point ${current.Point}.");
                    }

                    if (_chain.IsCurrentInCycle)
                    {
                        throw new InvalidOperationException($"Formula in a cell '${sheetInfo.Sheet.Name}'!${current.Point} is part of a cycle.");
                    }

                    var cellFormula = sheetInfo.FormulaSlice.Get(current.Point);
                    if (cellFormula is null)
                    {
                        throw new InvalidOperationException($"Calculation chain contains a '${sheetInfo.Sheet.Name}'!${current.Point}, but the cell doesn't contain formula.");
                    }

                    if (!cellFormula.IsDirty)
                        break;

                    try
                    {
                        ApplyFormula(cellFormula, current.Point, sheetInfo.Sheet, sheetInfo.ValueSlice,
                            recalculateSheetId);
                        cellFormula.IsDirty = false;

                        // Break out of the inner loop, a dirty cell has been
                        // calculated and thus chain can move ahead.
                        break;
                    }
                    catch (GettingDataException ex)
                    {
                        _chain.MoveToCurrent(ex.Point);
                    }
                }
            }

            // Super important to clean up the chain for next recalculation.
            // Chain contains shared data and not cleaning it would cause hard
            // to diagnose issues.
            _chain.Reset();
        }

        private void ApplyFormula(XLCellFormula formula, XLSheetPoint appliedPoint, XLWorksheet sheet, ValueSlice valueSlice, uint? recalculateSheetId)
        {
            var formulaText = formula.GetFormulaA1();
            if (formula.Type == FormulaType.Normal)
            {
                var single = EvaluateFormula(
                    formulaText,
                    sheet.Workbook,
                    sheet,
                    new XLAddress(sheet, appliedPoint.Row, appliedPoint.Column, true, true),
                    recalculateSheetId: recalculateSheetId);
                valueSlice.SetCellValue(appliedPoint, single.ToCellValue());
            }
            else if (formula.Type == FormulaType.Array)
            {
                // The point can be any point in an array, so we can't use it.
                var range = formula.Range;
                var leftTopCorner = range.FirstPoint;
                var masterCell = sheet.Cell(leftTopCorner.Row, leftTopCorner.Column);
                var array = EvaluateArrayFormula(formulaText, masterCell, recalculateSheetId);

                // The array from formula can be smaller or larger than the
                // range of cells it should fit into. Broadcast it to the size.
                var result = array.Broadcast(range.Height, range.Width);

                // Copy value to the value slice
                for (var rowIdx = 0; rowIdx < result.Height; ++rowIdx)
                {
                    for (var colIdx = 0; colIdx < result.Width; ++colIdx)
                    {
                        var cellValue = result[rowIdx, colIdx];
                        var row = range.FirstPoint.Row + rowIdx;
                        var column = range.FirstPoint.Column + colIdx;
                        valueSlice.SetCellValue(new XLSheetPoint(row, column), cellValue.ToCellValue());
                    }
                }
            }
            else
            {
                throw new NotImplementedException($"Evaluation of formula type '{formula.Type}' is not supported.");
            }
        }

        /// <summary>
        /// Evaluates a normal formula.
        /// </summary>
        /// <param name="expression">Expression to evaluate.</param>
        /// <param name="wb">Workbook where is formula being evaluated.</param>
        /// <param name="ws">Worksheet where is formula being evaluated.</param>
        /// <param name="address">Address of formula.</param>
        /// <param name="recursive">Should the data necessary for this formula (not deeper ones)
        /// be calculated recursively? Used only for non-cell calculations.</param>
        /// <param name="recalculateSheetId">
        /// If set, calculation  will allow dirty reads from other sheets than the passed one.
        /// </param>
        /// <returns>The value of the expression.</returns>
        /// <remarks>
        /// If you are going to evaluate the same expression several times,
        /// it is more efficient to parse it only once using the <see cref="Parse"/>
        /// method and then using the Expression.Evaluate method to evaluate
        /// the parsed expression.
        /// </remarks>
        internal ScalarValue EvaluateFormula(string expression, XLWorkbook? wb = null, XLWorksheet? ws = null, IXLAddress? address = null, bool recursive = false, uint? recalculateSheetId = null)
        {
            var ctx = new CalcContext(this, _culture, wb, ws, address, recursive)
            {
                RecalculateSheetId = recalculateSheetId
            };
            var result = EvaluateFormula(expression, ctx);
            if (ctx.UseImplicitIntersection)
            {
                result = result.Match(
                    () => AnyValue.Blank,
                    logical => logical,
                    number => number,
                    text => text,
                    error => error,
                    array => array[0, 0].ToAnyValue(),
                    reference => reference);
            }

            return ToCellContentValue(result, ctx);
        }

        private Array EvaluateArrayFormula(string expression, XLCell masterCell, uint? recalculateSheetId)
        {
            var ctx = new CalcContext(this, _culture, masterCell)
            {
                IsArrayCalculation = true,
                RecalculateSheetId = recalculateSheetId
            };
            var result = EvaluateFormula(expression, ctx);
            if (result.TryPickSingleOrMultiValue(out var single, out var multi, ctx))
                return new ScalarArray(single, 1, 1);

            return multi;
        }

        internal AnyValue EvaluateName(string nameFormula, XLWorksheet ws)
        {
            var ctx = new CalcContext(this, _culture, ws.Workbook, ws, null);
            return EvaluateFormula(nameFormula, ctx);
        }

        private AnyValue EvaluateFormula(string expression, CalcContext ctx)
        {
            var x = _cache[expression];

            var result = x.AstRoot.Accept(ctx, _visitor);
            return result;
        }

        // build/get static keyword table
        private FunctionRegistry GetFunctionTable()
        {
            var fr = new FunctionRegistry();

            // register built-in functions (and constants)
            Engineering.Register(fr);
            Information.Register(fr);
            Logical.Register(fr);
            Lookup.Register(fr);
            MathTrig.Register(fr);
            Text.Register(fr);
            Statistical.Register(fr);
            DateAndTime.Register(fr);
            Financial.Register(fr);

            return fr;
        }

        /// <summary>
        /// Convert any kind of formula value to value returned as a content of a cell.
        /// <list type="bullet">
        ///    <item><c>bool</c> - represents a logical value.</item>
        ///    <item><c>double</c> - represents a number and also date/time as serial date-time.</item>
        ///    <item><c>string</c> - represents a text value.</item>
        ///    <item><see cref="XLError" /> - represents a formula calculation error.</item>
        /// </list>
        /// </summary>
        private static ScalarValue ToCellContentValue(AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar;

            if (collection.TryPickT0(out var array, out var reference))
            {
                return array![0, 0];
            }

            if (reference!.TryGetSingleCellValue(out var cellValue, ctx))
                return cellValue;

            var intersected = reference.ImplicitIntersection(ctx.FormulaAddress);
            if (!intersected.TryPickT0(out var singleCellReference, out var error))
                return error;

            if (!singleCellReference!.TryGetSingleCellValue(out var singleCellValue, ctx))
                throw new InvalidOperationException("Got multi cell reference instead of single cell reference.");

            return singleCellValue;
        }

        void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        {
            if (_dependencyTree is not null)
                _dependencyTree.RenameSheet(oldSheetName, newSheetName);
        }
    }

    internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue> arg);

    /// <summary>
    /// Delegate that represents CalcEngine functions.
    /// </summary>
    /// <param name="parms">List of <see cref="Expression"/> objects that represent the
    /// parameters to be used in the function call.</param>
    /// <returns>The function result.</returns>
    internal delegate object LegacyCalcEngineFunction(List<Expression> parms);
}
