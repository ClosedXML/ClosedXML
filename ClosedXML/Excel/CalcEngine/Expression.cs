using ClosedXML.Excel.CalcEngine.Exceptions;
using ClosedXML.Excel.Patterns;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal abstract class ExpressionBase
    {
        public abstract string LastParseItem { get; }
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

        internal Expression()
        {
            _token = new Token(null, TKID.ATOM, TKTYPE.IDENTIFIER);
        }

        internal Expression(object value)
        {
            _token = new Token(value, TKID.ATOM, TKTYPE.LITERAL);
        }

        internal Expression(Token tk)
        {
            _token = tk;
        }

        #endregion ** ctors

        //---------------------------------------------------------------------------

        #region ** object model

        public virtual object Evaluate()
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
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            var v = new ConvertibleObject(x.Evaluate());
            return (string)v;
        }

        public static implicit operator double(Expression x)
        {
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            var v = new ConvertibleObject(x.Evaluate());
            return (double)v;
        }

        public static implicit operator bool(Expression x)
        {
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            var v = new ConvertibleObject(x.Evaluate());
            return (bool)v;
        }

        public static implicit operator DateTime(Expression x)
        {
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            var v = new ConvertibleObject(x.Evaluate());
            return (DateTime)v;
        }

        #endregion ** implicit converters

        //---------------------------------------------------------------------------

        #region ** IComparable<Expression>

        public int CompareTo(Expression other)
        {
            return ClosedXMLValueComparer.DefaultComparer.Compare(this, other);
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
        public UnaryExpression(Token tk, Expression expr) : base(tk)
        {
            _expr = expr;
        }

        // ** object model
        override public object Evaluate()
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
                ? new Expression(this.Evaluate())
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
        public BinaryExpression(Token tk, Expression exprLeft, Expression exprRight) : base(tk)
        {
            _lft = exprLeft;
            _rgt = exprRight;
        }

        // ** object model
        override public object Evaluate()
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
                    if (Math.Abs((double)_rgt) < double.Epsilon)
                        throw new DivisionByZeroException();

                    return (double)_lft / (double)_rgt;

                case TKID.DIVINT:
                    if (Math.Abs((double)_rgt) < double.Epsilon)
                        throw new DivisionByZeroException();

                    return (double)(int)((double)_lft / (double)_rgt);

                case TKID.MOD:
                    if (Math.Abs((double)_rgt) < double.Epsilon)
                        throw new DivisionByZeroException();

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
                ? new Expression(this.Evaluate())
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
        internal FunctionExpression()
        { }

        public FunctionExpression(FunctionDefinition function, List<Expression> parms)
        {
            _fn = function;
            _parms = parms;
        }

        // ** object model
        override public object Evaluate()
        {
            return _fn.Function(_parms);
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
                ? new Expression(this.Evaluate())
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

        public VariableExpression(Dictionary<string, object> dct, string name)
        {
            _dct = dct;
            _name = name;
        }

        public override object Evaluate()
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
        internal XObjectExpression(object value)
        {
            _value = value;
        }

        public object Value { get { return _value; } }

        // ** object model
        public override object Evaluate()
        {
            // use IValueObject if available
            var iv = _value as IValueObject;
            if (iv != null)
            {
                return iv.GetValue();
            }

            // return raw object
            return _value;
        }

        public IEnumerator GetEnumerator()
        {
            if (_value is string)
                return new[] { (string)_value }.GetEnumerator();

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
        internal EmptyValueExpression() { }

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

        internal ErrorExpression(ExpressionErrorType eet)
            : base(new Token(eet, TKID.ATOM, TKTYPE.ERROR))
        { }

        public override object Evaluate()
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
