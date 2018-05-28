using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace ClosedXML.Excel.CalcEngine
{
    internal abstract class ExpressionBase
    {
        private CalculationContext _context;
        protected ExpressionBase(in CalculationContext ctx)
        {
            _context = ctx;
        }

        public abstract string LastParseItem { get; }

        public ref CalculationContext CurrentContext
        {
            get => ref _context;
        }
    }

    /// <summary>
    /// Base class that represents parsed expressions.
    /// </summary>
    /// <remarks>
    /// For example:
    /// <code>
    /// Expression expr = scriptEngine.Parse(strExpression);
    /// object val = expr.Evaluate();
    /// </code>
    /// </remarks>
    internal class Expression : ExpressionBase, IComparable<Expression>
    {
        //---------------------------------------------------------------------------

        #region ** fields

        internal readonly Token _token;

        #endregion ** fields

        //---------------------------------------------------------------------------

        #region ** ctors

        internal Expression(in CalculationContext ctx)
            : base(in ctx)
        {
            _token = new Token(null, TKID.ATOM, TKTYPE.IDENTIFIER);
        }

        internal Expression(in CalculationContext ctx, object value)
            : base(in ctx)
        {
            _token = new Token(value, TKID.ATOM, TKTYPE.LITERAL);
        }

        internal Expression(in CalculationContext ctx, Token tk)
            : base(in ctx)
        {
            _token = tk;
        }

        #endregion ** ctors

        //---------------------------------------------------------------------------

        #region ** object model

        public virtual object Evaluate()
        {
            return Evaluate(in this.CurrentContext);
        }
        
        public virtual object Evaluate(in CalculationContext ctx)
        {
            if (_token.Type != TKTYPE.LITERAL)
            {
                throw new ArgumentException("Bad expression.");
            }
            return _token.Value;
        }

        public virtual Expression Optimize()
        {
            return this;
        }

        #endregion ** object model

        //---------------------------------------------------------------------------

        #region ** implicit converters

        public static implicit operator string(Expression x)
        {
            if (x is ErrorExpression ee)
                ee.ThrowApplicableException();

            var v = x.Evaluate(in x.CurrentContext);
            return v == null ? string.Empty : v.ToString();
        }

        public static implicit operator double(Expression x)
        {
            if (x is ErrorExpression ee)
                ee.ThrowApplicableException();

            // evaluate
            var v = x.Evaluate(in x.CurrentContext);

            // handle doubles
            if (v is double)
            {
                return (double)v;
            }

            // handle booleans
            if (v is bool)
            {
                return (bool)v ? 1 : 0;
            }

            // handle dates
            if (v is DateTime)
            {
                return ((DateTime)v).ToOADate();
            }

            if (v is TimeSpan)
            {
                return ((TimeSpan)v).TotalDays;
            }

            // handle nulls
            if (v == null || v is string)
            {
                return 0;
            }

            // handle everything else
            CultureInfo _ci = Thread.CurrentThread.CurrentCulture;
            return (double)Convert.ChangeType(v, typeof(double), _ci);
        }

        public static implicit operator bool(Expression x)
        {
            if (x is ErrorExpression ee)
                ee.ThrowApplicableException();

            // evaluate
            var v = x.Evaluate(x.CurrentContext);

            // handle booleans
            if (v is bool)
            {
                return (bool)v;
            }

            // handle nulls
            if (v == null)
            {
                return false;
            }

            // handle doubles
            if (v is double)
            {
                return (double)v == 0 ? false : true;
            }

            // handle everything else
            return (double)x == 0 ? false : true;
        }

        public static implicit operator DateTime(Expression x)
        {
            if (x is ErrorExpression ee)
                ee.ThrowApplicableException();

            // evaluate
            var v = x.Evaluate(x.CurrentContext);

            // handle dates
            if (v is DateTime)
            {
                return (DateTime)v;
            }

            // handle doubles
            if (v is double || v is int)
            {
                return DateTime.FromOADate((double)x);
            }

            // handle everything else
            CultureInfo _ci = Thread.CurrentThread.CurrentCulture;
            return (DateTime)Convert.ChangeType(v, typeof(DateTime), _ci);
        }

        #endregion ** implicit converters

        //---------------------------------------------------------------------------

        #region ** IComparable<Expression>

        public int CompareTo(Expression other)
        {
            // get both values
            var c1 = this.Evaluate(in this.CurrentContext) as IComparable;
            var c2 = other.Evaluate(in other.CurrentContext) as IComparable;

            // handle nulls
            if (c1 == null && c2 == null)
            {
                return 0;
            }
            if (c2 == null)
            {
                return -1;
            }
            if (c1 == null)
            {
                return +1;
            }

            // make sure types are the same
            if (c1.GetType() != c2.GetType())
            {
                try
                {
                    if (c1 is DateTime)
                        c2 = ((DateTime)other);
                    else if (c2 is DateTime)
                        c1 = ((DateTime)this);
                    else
                        c2 = Convert.ChangeType(c2, c1.GetType()) as IComparable;
                }
                catch (InvalidCastException) { return -1; }
                catch (FormatException) { return -1; }
                catch (OverflowException) { return -1; }
                catch (ArgumentNullException) { return -1; }
            }

            // compare
            return c1.CompareTo(c2);
        }

        #endregion ** IComparable<Expression>

        //---------------------------------------------------------------------------

        #region ** ExpressionBase

        public override string LastParseItem
        {
            get { return _token?.Value?.ToString() ?? "Unknown value"; }
        }

        #endregion ** ExpressionBase
    }

    /// <summary>
    /// Unary expression, e.g. +123
    /// </summary>
    internal class UnaryExpression : Expression
    {
        // ** fields
        private Expression _expr;

        // ** ctor
        public UnaryExpression(in CalculationContext ctx, Token tk, Expression expr)
            : base(in ctx, tk)
        {
            _expr = expr;
        }

        // ** object model
        override public object Evaluate(in CalculationContext ctx)
        {
            switch (_token.ID)
            {
                case TKID.ADD:
                    return +(double)_expr;

                case TKID.SUB:
                    return -(double)_expr;
            }
            throw new ArgumentException("Bad expression.");
        }

        public override Expression Optimize()
        {
            _expr = _expr.Optimize();
            return _expr._token.Type == TKTYPE.LITERAL
                ? new Expression(in this.CurrentContext, this.Evaluate(in this.CurrentContext))
                : this;
        }

        public override string LastParseItem
        {
            get { return _expr.LastParseItem; }
        }
    }

    /// <summary>
    /// Binary expression, e.g. 1+2
    /// </summary>
    internal class BinaryExpression : Expression
    {
        // ** fields
        private Expression _lft;

        private Expression _rgt;

        // ** ctor
        public BinaryExpression(in CalculationContext ctx, Token tk, Expression exprLeft, Expression exprRight)
            : base(in ctx, tk)
        {
            _lft = exprLeft;
            _rgt = exprRight;
        }

        // ** object model
        override public object Evaluate(in CalculationContext ctx)
        {
            // handle comparisons
            if (_token.Type == TKTYPE.COMPARE)
            {
                var cmp = _lft.CompareTo(_rgt);
                switch (_token.ID)
                {
                    case TKID.GT: return cmp > 0;
                    case TKID.LT: return cmp < 0;
                    case TKID.GE: return cmp >= 0;
                    case TKID.LE: return cmp <= 0;
                    case TKID.EQ: return cmp == 0;
                    case TKID.NE: return cmp != 0;
                }
            }

            // handle everything else
            switch (_token.ID)
            {
                case TKID.CONCAT:
                    return (string)_lft + (string)_rgt;

                case TKID.ADD:
                    return (double)_lft + (double)_rgt;

                case TKID.SUB:
                    return (double)_lft - (double)_rgt;

                case TKID.MUL:
                    return (double)_lft * (double)_rgt;

                case TKID.DIV:
                    return (double)_lft / (double)_rgt;

                case TKID.DIVINT:
                    return (double)(int)((double)_lft / (double)_rgt);

                case TKID.MOD:
                    return (double)(int)((double)_lft % (double)_rgt);

                case TKID.POWER:
                    var a = (double)_lft;
                    var b = (double)_rgt;
                    if (b == 0.0) return 1.0;
                    if (b == 0.5) return Math.Sqrt(a);
                    if (b == 1.0) return a;
                    if (b == 2.0) return a * a;
                    if (b == 3.0) return a * a * a;
                    if (b == 4.0) return a * a * a * a;
                    return Math.Pow((double)_lft, (double)_rgt);
            }
            throw new ArgumentException("Bad expression.");
        }

        public override Expression Optimize()
        {
            _lft = _lft.Optimize();
            _rgt = _rgt.Optimize();
            return _lft._token.Type == TKTYPE.LITERAL && _rgt._token.Type == TKTYPE.LITERAL
                ? new Expression(in this.CurrentContext, this.Evaluate(in this.CurrentContext))
                : this;
        }

        public override string LastParseItem
        {
            get { return _rgt.LastParseItem; }
        }
    }

    /// <summary>
    /// Function call expression, e.g. sin(0.5)
    /// </summary>
    internal class FunctionExpression : Expression
    {
        // ** fields
        private readonly FunctionDefinition _fn;

        private readonly List<Expression> _parms;

        // ** ctor
        internal FunctionExpression(in CalculationContext ctx)
            : base(in ctx)
        { }

        public FunctionExpression(in CalculationContext ctx, FunctionDefinition function, List<Expression> parms)
            : base(in ctx)
        {
            _fn = function;
            _parms = parms;
        }

        // ** object model
        override public object Evaluate(in CalculationContext ctx)
        {
            return _fn.Function(in ctx, _parms);
        }

        public override Expression Optimize()
        {
            bool allLits = true;
            if (_parms != null)
            {
                for (int i = 0; i < _parms.Count; i++)
                {
                    var p = _parms[i].Optimize();
                    _parms[i] = p;
                    if (p._token.Type != TKTYPE.LITERAL)
                    {
                        allLits = false;
                    }
                }
            }
            return allLits
                ? new Expression(in this.CurrentContext, this.Evaluate(in this.CurrentContext))
                : this;
        }

        public override string LastParseItem
        {
            get { return _parms.Last().LastParseItem; }
        }
    }

    /// <summary>
    /// Simple variable reference.
    /// </summary>
    internal class VariableExpression : Expression
    {
        private readonly Dictionary<string, object> _dct;
        private readonly string _name;

        public VariableExpression(in CalculationContext ctx, Dictionary<string, object> dct, string name)
            : base(in ctx)
        {
            _dct = dct;
            _name = name;
        }

        public override object Evaluate(in CalculationContext ctx)
        {
            return _dct[_name];
        }

        public override string LastParseItem
        {
            get { return _name; }
        }
    }

    /// <summary>
    /// Expression that represents an external object.
    /// </summary>
    internal class XObjectExpression :
        Expression,
        IEnumerable
    {
        private readonly object _value;

        // ** ctor
        internal XObjectExpression(in CalculationContext ctx, object value)
            : base(in ctx)
        {
            _value = value;
        }

        public object Value { get { return _value; } }

        // ** object model
        public override object Evaluate(in CalculationContext ctx)
        {
            // use IValueObject if available
            if (_value is IValueObject iv)
            {
                return iv.GetValue();
            }

            // return raw object
            return _value;
        }

        public IEnumerator GetEnumerator()
        {
            if (_value is string s)
                return new[] { s }.GetEnumerator();

            return (_value as IEnumerable).GetEnumerator();
        }

        public override string LastParseItem
        {
            get { return Value.ToString(); }
        }
    }

    /// <summary>
    /// Expression that represents an omitted parameter.
    /// </summary>
    internal class EmptyValueExpression : Expression
    {
        internal EmptyValueExpression(in CalculationContext ctx)
        : base(in ctx)
        { }

        public override string LastParseItem
        {
            get { return "<EMPTY VALUE>"; }
        }
    }

    internal class ErrorExpression : Expression
    {
        internal enum ExpressionErrorType
        {
            CellReference,
            CellValue,
            DivisionByZero,
            NameNotRecognized,
            NoValueAvailable,
            NullValue,
            NumberInvalid
        }

        internal ErrorExpression(in CalculationContext ctx, ExpressionErrorType eet)
            : base(in ctx, new Token(eet, TKID.ATOM, TKTYPE.ERROR))
        { }

        public override object Evaluate(in CalculationContext ctx)
        {
            return this._token.Value;
        }

        public void ThrowApplicableException()
        {
            var eet = (ExpressionErrorType)_token.Value;
            switch (eet)
            {
                // TODO: include last token in exception message
                case ExpressionErrorType.CellReference:
                    throw new CellReferenceException();
                case ExpressionErrorType.CellValue:
                    throw new CellValueException();
                case ExpressionErrorType.DivisionByZero:
                    throw new DivisionByZeroException();
                case ExpressionErrorType.NameNotRecognized:
                    throw new NameNotRecognizedException();
                case ExpressionErrorType.NoValueAvailable:
                    throw new NoValueAvailableException();
                case ExpressionErrorType.NullValue:
                    throw new NullValueException();
                case ExpressionErrorType.NumberInvalid:
                    throw new NumberException();
            }
        }
    }

    /// <summary>
    /// Interface supported by external objects that have to return a value
    /// other than themselves (e.g. a cell range object should return the
    /// cell content instead of the range itself).
    /// </summary>
    public interface IValueObject
    {
        object GetValue();
    }
}
