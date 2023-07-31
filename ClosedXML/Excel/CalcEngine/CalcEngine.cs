using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;
using System.Globalization;

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
    internal class CalcEngine
    {
        private readonly CultureInfo _culture;
        private readonly ExpressionCache _cache;               // cache with parsed expressions
        private readonly FormulaParser _parser;
        private readonly CalculationVisitor _visitor;

        public CalcEngine(CultureInfo culture)
        {
            _culture = culture;
            _cache = new ExpressionCache(this);
            var funcRegistry = GetFunctionTable();
            _parser = new FormulaParser(funcRegistry);
            _visitor = new CalculationVisitor(funcRegistry);
        }

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public Formula Parse(string expression)
        {
            return _parser.GetAst(expression);
        }

        /// <summary>
        /// Evaluates a normal formula.
        /// </summary>
        /// <param name="expression">Expression to evaluate.</param>
        /// <param name="wb">Workbook where is formula being evaluated.</param>
        /// <param name="ws">Worksheet where is formula being evaluated.</param>
        /// <param name="address">Address of formula.</param>
        /// <returns>The value of the expression.</returns>
        /// <remarks>
        /// If you are going to evaluate the same expression several times,
        /// it is more efficient to parse it only once using the <see cref="Parse"/>
        /// method and then using the Expression.Evaluate method to evaluate
        /// the parsed expression.
        /// </remarks>
        internal ScalarValue EvaluateFormula(string expression, XLWorkbook? wb = null, XLWorksheet? ws = null, IXLAddress? address = null)
        {
            var ctx = new CalcContext(this, _culture, wb, ws, address);
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

        internal Array EvaluateArrayFormula(string expression, XLCell masterCell)
        {
            var ctx = new CalcContext(this, _culture, masterCell) { IsArrayCalculation = true };
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
