using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;
using System.Globalization;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

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
        protected readonly CultureInfo _culture;
        protected ExpressionCache _cache;               // cache with parsed expressions
        private readonly FormulaParser _parser;
        private readonly FunctionRegistry _funcRegistry;      // table with constants and functions (pi, sin, etc)

        public CalcEngine()
        {
            _funcRegistry = GetFunctionTable();
            _cache = new ExpressionCache(this);
            _parser = new FormulaParser(_funcRegistry);
        }

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public ValueNode Parse(string expression)
        {
            var cstTree = _parser.Parse(expression);
            var root = (ValueNode)cstTree.Root.AstNode ?? throw new InvalidOperationException("Formula doesn't have AST root.");
            return root;//(Expression)root.Accept(null, _compatibilityVisitor);
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

            var ctx = new CalcContext(_culture, wb, ws, address);
            var calculatingVisitor = new CalculationVisitor(_funcRegistry);
            var result = x.Accept(ctx, calculatingVisitor);
            if (ctx.UseImplicitIntersection && result.IsT4)
            {
                result = result.AsT4[0, 0].ToAnyValue();
            }

            // TODO exception
            return result.Match<object>(logical => logical.Value,
                number => number.Value,
                text => text.Value,
                error => error,
                array => throw new InvalidOperationException("Array shouldn't be present currently"),
                reference => throw new NotImplementedException("WTF with this?")); 

            //return x.Evaluate();
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

        /// <summary>
        /// Gets an external object based on an identifier.
        /// </summary>
        /// <remarks>
        /// This method is useful when the engine needs to create objects dynamically.
        /// For example, a spreadsheet calc engine would use this method to dynamically create cell
        /// range objects based on identifiers that cannot be enumerated at design time
        /// (such as "AB12", "A1:AB12", etc.)
        /// </remarks>
        public virtual object GetExternalObject(string identifier)
        {
            return null;
        }

        // build/get static keyword table
        private FunctionRegistry GetFunctionTable()
        {
            var fr = new FunctionRegistry();

            // register built-in functions (and constants)
            Engineering.Register(fr);
            Information.Register(fr);
            LogicalFunctions.Register(fr);
            Lookup.Register(fr);
            MathTrig.Register(fr);
            TextFunctions.Register(fr);
            Statistical.Register(fr);
            DateAndTime.Register(fr);
            Financial.Register(fr);

            return fr;
        }
    }

    internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue?> arg);

    /// <summary>
    /// Delegate that represents CalcEngine functions.
    /// </summary>
    /// <param name="parms">List of <see cref="Expression"/> objects that represent the
    /// parameters to be used in the function call.</param>
    /// <returns>The function result.</returns>
    internal delegate object LegacyCalcEngineFunction(List<Expression> parms);
}
