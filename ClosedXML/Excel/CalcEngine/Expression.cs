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

            var v = x.Evaluate();

            if (v == null)
                return string.Empty;

            if (v is bool b)
                return b.ToString().ToUpper();

            return v.ToString();
        }

        public static implicit operator double(Expression x)
        {
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            // evaluate
            var v = x.Evaluate();

            // handle doubles
            if (v is double dbl)
            {
                return dbl;
            }

            // handle booleans
            if (v is bool b)
            {
                return b ? 1 : 0;
            }

            // handle dates
            if (v is DateTime dt)
            {
                return dt.ToOADate();
            }

            if (v is TimeSpan ts)
            {
                return ts.TotalDays;
            }

            // handle string
            if (v is string s && double.TryParse(s, out var doubleValue))
            {
                return doubleValue;
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
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            // evaluate
            var v = x.Evaluate();

            // handle booleans
            if (v is bool b)
            {
                return b;
            }

            // handle nulls
            if (v == null)
            {
                return false;
            }

            // handle doubles
            if (v is double dbl)
            {
                return dbl != 0;
            }

            // handle everything else
            return (double)Convert.ChangeType(v, typeof(double)) != 0;
        }

        public static implicit operator DateTime(Expression x)
        {
            if (x is ErrorExpression)
                (x as ErrorExpression).ThrowApplicableException();

            // evaluate
            var v = x.Evaluate();

            // handle dates
            if (v is DateTime dt)
            {
                return dt;
            }

            if (v is TimeSpan ts)
            {
                return new DateTime().Add(ts);
            }

            // handle numbers
            if (v.IsNumber())
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
            var c1 = this.Evaluate() as IComparable;
            var c2 = other.Evaluate() as IComparable;

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

            // String comparisons should be case insensitive
            if (c1 is string s1 && c2 is string s2)
                return StringComparer.OrdinalIgnoreCase.Compare(s1, s2);
            else
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
        // ** ctor
        public UnaryExpression(Token tk, Expression expr) : base(tk)
        {
            Expression = expr;
        }

        public Expression Expression { get; private set; }

        // ** object model
        override public object Evaluate()
        {
            switch (_token.ID)
            {
                case TKID.ADD:
                    return +(double)Expression;

                case TKID.SUB:
                    return -(double)Expression;
            }
            throw new ArgumentException("Bad expression.");
        }

        public override Expression Optimize()
        {
            Expression = Expression.Optimize();
            return Expression._token.Type == TKTYPE.LITERAL
                ? new Expression(this.Evaluate())
                : this;
        }

        public override string LastParseItem
        {
            get { return Expression.LastParseItem; }
        }
    }

    /// <summary>
    /// Binary expression, e.g. 1+2
    /// </summary>
    internal class BinaryExpression : Expression
    {
        // ** ctor
        public BinaryExpression(Token tk, Expression exprLeft, Expression exprRight) : base(tk)
        {
            LeftExpression = exprLeft;
            RightExpression = exprRight;
        }

        public Expression LeftExpression { get; private set; }
        public Expression RightExpression { get; private set; }

        // ** object model
        override public object Evaluate()
        {
            // handle comparisons
            if (_token.Type == TKTYPE.COMPARE)
            {
                var cmp = LeftExpression.CompareTo(RightExpression);
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
                    return (string)LeftExpression + (string)RightExpression;

                case TKID.ADD:
                    return (double)LeftExpression + (double)RightExpression;

                case TKID.SUB:
                    return (double)LeftExpression - (double)RightExpression;

                case TKID.MUL:
                    return (double)LeftExpression * (double)RightExpression;

                case TKID.DIV:
                    if (Math.Abs((double)RightExpression) < double.Epsilon)
                        return XLCalculationErrorType.DivisionByZero;

                    return (double)LeftExpression / (double)RightExpression;

                case TKID.DIVINT:
                    if (Math.Abs((double)RightExpression) < double.Epsilon)
                        return XLCalculationErrorType.DivisionByZero;

                    return (double)(int)((double)LeftExpression / (double)RightExpression);

                case TKID.MOD:
                    if (Math.Abs((double)RightExpression) < double.Epsilon)
                        return XLCalculationErrorType.DivisionByZero;

                    return (double)(int)((double)LeftExpression % (double)RightExpression);

                case TKID.POWER:
                    var a = (double)LeftExpression;
                    var b = (double)RightExpression;
                    if (b == 0.0) return 1.0;
                    if (b == 0.5) return Math.Sqrt(a);
                    if (b == 1.0) return a;
                    if (b == 2.0) return a * a;
                    if (b == 3.0) return a * a * a;
                    if (b == 4.0) return a * a * a * a;
                    return Math.Pow((double)LeftExpression, (double)RightExpression);
            }
            throw new ArgumentException("Bad expression.");
        }

        public override Expression Optimize()
        {
            LeftExpression = LeftExpression.Optimize();
            RightExpression = RightExpression.Optimize();
            return LeftExpression._token.Type == TKTYPE.LITERAL && RightExpression._token.Type == TKTYPE.LITERAL
                ? new Expression(this.Evaluate())
                : this;
        }

        public override string LastParseItem
        {
            get { return RightExpression.LastParseItem; }
        }
    }

    /// <summary>
    /// Function call expression, e.g. sin(0.5)
    /// </summary>
    internal class FunctionExpression : Expression
    {
        // ** ctor
        internal FunctionExpression()
        { }

        public FunctionExpression(FunctionDefinition function, List<Expression> parms)
        {
            FunctionDefinition = function;
            Parameters = parms;
        }

        // ** object model
        override public object Evaluate()
        {
            if (FunctionDefinition.EvaluateParameters && Parameters != null)
            {
                foreach (var p in Parameters)
                {
                    if (p is ErrorExpression errorExpression)
                        return errorExpression.Evaluate();
                }
            }

            try
            {
                return FunctionDefinition.Function(Parameters);
            }
            catch (CellReferenceException)
            {
                return XLCalculationErrorType.CellReference;
            }
            catch (CellValueException)
            {
                return XLCalculationErrorType.CellValue;
            }
            catch (DivisionByZeroException)
            {
                return XLCalculationErrorType.DivisionByZero;
            }
            catch (NameNotRecognizedException)
            {
                return XLCalculationErrorType.NameNotRecognized;
            }
            catch (NoValueAvailableException)
            {
                return XLCalculationErrorType.NoValueAvailable;
            }
            catch (NullValueException)
            {
                return XLCalculationErrorType.NullValue;
            }
            catch (NumberException)
            {
                return XLCalculationErrorType.NumberInvalid;
            }
        }

        public FunctionDefinition FunctionDefinition { get; }
        public List<Expression> Parameters { get; }

        public override Expression Optimize()
        {
            bool allLits = true;
            if (Parameters != null)
            {
                for (int i = 0; i < Parameters.Count; i++)
                {
                    var p = Parameters[i].Optimize();
                    Parameters[i] = p;
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
            get { return Parameters.Last().LastParseItem; }
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
    internal class XObjectExpression : Expression, IEnumerable
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
            if (_value is string s)
            {
                yield return s;
            }
            else if (_value is IEnumerable ie)
            {
                foreach (var o in ie)
                    yield return o;
            }
            else
            {
                yield return _value;
            }
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
        internal EmptyValueExpression()
            // Ensures a token of type LITERAL, with value of null is created
            : base(value: null) 
        {
        }

        public override string LastParseItem
        {
            get { return "<EMPTY VALUE>"; }
        }
    }

    internal class ErrorExpression : Expression
    {
        internal ErrorExpression(XLCalculationErrorType eet)
            : base(new Token(eet, TKID.ATOM, TKTYPE.ERROR))
        { }

        public override object Evaluate()
        {
            return this._token.Value;
        }

        // To be used only in implicit operators
        public void ThrowApplicableException()
        {
            var eet = (XLCalculationErrorType)_token.Value;
            switch (eet)
            {
                // TODO: include last token in exception message
                case XLCalculationErrorType.CellReference:
                    throw new CellReferenceException();
                case XLCalculationErrorType.CellValue:
                    throw new CellValueException();
                case XLCalculationErrorType.DivisionByZero:
                    throw new DivisionByZeroException();
                case XLCalculationErrorType.NameNotRecognized:
                    throw new NameNotRecognizedException();
                case XLCalculationErrorType.NoValueAvailable:
                    throw new NoValueAvailableException();
                case XLCalculationErrorType.NullValue:
                    throw new NullValueException();
                case XLCalculationErrorType.NumberInvalid:
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
