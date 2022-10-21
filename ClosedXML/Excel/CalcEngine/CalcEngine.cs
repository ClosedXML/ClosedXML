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
        private ExpressionCache _cache;               // cache with parsed expressions
        private readonly FormulaParser _parser;
        private readonly FunctionRegistry _funcRegistry;      // table with constants and functions (pi, sin, etc)
        private readonly CalculationVisitor _visitor;

        public CalcEngine(CultureInfo culture)
        {
            _culture = culture;
            _funcRegistry = GetFunctionTable();
            _cache = new ExpressionCache(this);
            _parser = new FormulaParser(_funcRegistry);
            _visitor = new CalculationVisitor(_funcRegistry);
        }

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public Formula Parse(string expression)
        {
            var cst = _parser.ParseCst(expression);
            return _parser.ConvertToAst(cst);
        }

        /// <summary>
        /// Evaluates an expression.
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
        public object Evaluate(string expression, XLWorkbook wb = null, XLWorksheet ws = null, IXLAddress address = null)
        {
            var x = _cache != null
                ? _cache[expression]
                : Parse(expression);

            var ctx = new CalcContext(this, _culture, wb, ws, address);
            var result = x.AstRoot.Accept(ctx, _visitor);
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

        internal AnyValue EvaluateExpression(string expression, XLWorkbook wb = null, XLWorksheet ws = null, IXLAddress address = null)
        {
            // Yay, copy pasta.
            var x = _cache != null
                    ? _cache[expression]
                    : Parse(expression);

            var ctx = new CalcContext(this, _culture, wb, ws, address);
            var calculatingVisitor = new CalculationVisitor(_funcRegistry);
            return x.AstRoot.Accept(ctx, calculatingVisitor);
        }

        /// <summary>
        /// Gets or sets whether the calc engine should keep a cache with parsed
        /// expressions.
        /// </summary>
        public bool CacheExpressions
        {
            get { return _cache != null; }
            set
            {
                if (value != CacheExpressions)
                {
                    _cache = value
                        ? new ExpressionCache(this)
                        : null;
                }
            }
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
        private static object ToCellContentValue(AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return ToCellContentValue(scalar);

            if (collection.TryPickT0(out var array, out var reference))
            {
                return ToCellContentValue(array[0, 0]);
            }

            if (reference.TryGetSingleCellValue(out var cellValue, ctx))
                return ToCellContentValue(cellValue);

            var intersected = reference.ImplicitIntersection(ctx.FormulaAddress);
            if (!intersected.TryPickT0(out var singleCellReference, out var error))
                return error;

            if (!singleCellReference.TryGetSingleCellValue(out var singleCellValue, ctx))
                throw new InvalidOperationException("Got multi cell reference instead of single cell reference.");

            return ToCellContentValue(singleCellValue);
        }

        private static object ToCellContentValue(ScalarValue value)
        {
            return value.Match<object>(
                () => 0,
                logical => logical,
                number => number,
                text => text,
                error => error);
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
