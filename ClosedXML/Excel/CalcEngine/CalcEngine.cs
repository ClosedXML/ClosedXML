using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;

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
        protected ExpressionCache _cache;               // cache with parsed expressions
        private readonly FormulaParser _parser;
        private readonly CompatibilityFormulaVisitor _compatibilityVisitor;
        private FunctionRegistry _fnTbl;      // table with constants and functions (pi, sin, etc)

        public CalcEngine()
        {
            _fnTbl = GetFunctionTable();
            _cache = new ExpressionCache(this);
            _parser = new FormulaParser(_fnTbl);
            _compatibilityVisitor = new CompatibilityFormulaVisitor(this);
        }

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public Expression Parse(string expression)
        {
            var cstTree = _parser.Parse(expression);
            var root = (Expression)cstTree.Root.AstNode ?? throw new InvalidOperationException("Formula doesn't have AST root.");
            return (Expression)root.Accept(null, _compatibilityVisitor);
        }

        /// <summary>
        /// Evaluates an expression.
        /// </summary>
        /// <param name="expression">Expression to evaluate.</param>
        /// <returns>The value of the expression.</returns>
        /// <remarks>
        /// If you are going to evaluate the same expression several times,
        /// it is more efficient to parse it only once using the <see cref="Parse"/>
        /// method and then using the Expression.Evaluate method to evaluate
        /// the parsed expression.
        /// </remarks>
        public object Evaluate(string expression)
        {
            var x = _cache != null
                    ? _cache[expression]
                    : Parse(expression);
            return x.Evaluate();
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
            if (_fnTbl == null)
            {
                // create table
                _fnTbl = new FunctionRegistry();

                // register built-in functions (and constants)
                Engineering.Register(_fnTbl);
                Information.Register(_fnTbl);
                Logical.Register(_fnTbl);
                Lookup.Register(_fnTbl);
                MathTrig.Register(_fnTbl);
                Text.Register(_fnTbl);
                Statistical.Register(_fnTbl);
                DateAndTime.Register(_fnTbl);
                Financial.Register(_fnTbl);
            }
            return _fnTbl;
        }
    }

    /// <summary>
    /// Delegate that represents CalcEngine functions.
    /// </summary>
    /// <param name="parms">List of <see cref="Expression"/> objects that represent the
    /// parameters to be used in the function call.</param>
    /// <returns>The function result.</returns>
    internal delegate object LegacyCalcEngineFunction(List<Expression> parms);
}
