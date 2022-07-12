using ClosedXML.Excel.CalcEngine.Exceptions;
using Irony.Ast;
using Irony.Parsing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using XLParser;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Base class for all AST nodes.
    /// </summary>
    internal abstract class ExpressionBase
    {
    }

    /// <summary>
    /// A node of AST that can be evaluated and produce a value.
    /// </summary>
    internal class Expression : ExpressionBase, IComparable<Expression>
    {
        internal Token _token;

        public Expression()
        {
            _token = new Token(null, TKTYPE.IDENTIFIER);
        }

        internal Expression(object value)
        {
            _token = new Token(value, TKTYPE.LITERAL);
        }

        public virtual TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

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
    }

    /// <summary>
    /// Unary expression, e.g. +123
    /// </summary>
    internal class UnaryExpression : Expression
    {
        public UnaryExpression(string operation, Expression expr) : this(null, operation, expr)
        { }

        public UnaryExpression(PrefixNode prefix, string operation, Expression expr)
        {
            Prefix = prefix;
            Operation = operation;
            Expression = expr;
        }

        public PrefixNode Prefix { get; }

        public string Operation { get; }

        public Expression Expression { get; private set; }

        // ** object model
        override public object Evaluate()
        {
            switch (Operation)
            {
                case "+":
                    return Expression.Evaluate();

                case "-":
                    return -(double)Expression;

                case "%":
                    return ((double)Expression) / 100.0;

                case "#":
                    throw new NotImplementedException("Evaluation of spill range operator is not implemented.");
            }
            throw new ArgumentException("Bad expression.");
        }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    internal enum BinaryOp
    {
        // Text operators
        Concat,
        // Arithmetic
        Add,
        Sub,
        Mult,
        Div,
        Exp,
        // Comparison operators
        Lt,
        Lte,
        Eq,
        Neq,
        Gte,
        Gt,
        // References operators
        Range,
        Union,
        Intersection
    }

    /// <summary>
    /// Binary expression, e.g. 1+2
    /// </summary>
    internal class BinaryExpression : Expression
    {
        private static readonly HashSet<BinaryOp> _comparisons = new HashSet<BinaryOp>
        {
            BinaryOp.Lt,
            BinaryOp.Lte,
            BinaryOp.Eq,
            BinaryOp.Neq,
            BinaryOp.Gte,
            BinaryOp.Gt
        };

        private readonly bool _isComparison;

        public BinaryExpression(BinaryOp operation, Expression exprLeft, Expression exprRight)
        {
            _isComparison = _comparisons.Contains(operation);
            Operation = operation;
            LeftExpression = exprLeft;
            RightExpression = exprRight;
        }

        public BinaryOp Operation { get; }

        public Expression LeftExpression { get; private set; }
        public Expression RightExpression { get; private set; }

        // ** object model
        override public object Evaluate()
        {
            // handle comparisons
            if (_isComparison)
            {
                var cmp = LeftExpression.CompareTo(RightExpression);
                switch (Operation)
                {
                    case BinaryOp.Gt: return cmp > 0;
                    case BinaryOp.Lt: return cmp < 0;
                    case BinaryOp.Gte: return cmp >= 0;
                    case BinaryOp.Lte: return cmp <= 0;
                    case BinaryOp.Eq: return cmp == 0;
                    case BinaryOp.Neq: return cmp != 0;
                }
            }

            // handle everything else
            switch (Operation)
            {
                case BinaryOp.Concat:
                    return (string)LeftExpression + (string)RightExpression;

                case BinaryOp.Add:
                    return (double)LeftExpression + (double)RightExpression;

                case BinaryOp.Sub:
                    return (double)LeftExpression - (double)RightExpression;

                case BinaryOp.Mult:
                    return (double)LeftExpression * (double)RightExpression;

                case BinaryOp.Div:
                    if (Math.Abs((double)RightExpression) < double.Epsilon)
                        throw new DivisionByZeroException();

                    return (double)LeftExpression / (double)RightExpression;

                case BinaryOp.Exp:
                    var a = (double)LeftExpression;
                    var b = (double)RightExpression;
                    if (b == 0.0) return 1.0;
                    if (b == 0.5) return Math.Sqrt(a);
                    if (b == 1.0) return a;
                    if (b == 2.0) return a * a;
                    if (b == 3.0) return a * a * a;
                    if (b == 4.0) return a * a * a * a;
                    return Math.Pow((double)LeftExpression, (double)RightExpression);
                case BinaryOp.Range:
                    throw new NotImplementedException("Evaluation of binary range operator is not implemented.");
                case BinaryOp.Union:
                    throw new NotImplementedException("Evaluation of range union operator is not implemented.");
                case BinaryOp.Intersection:
                    throw new NotImplementedException("Evaluation of range intersection operator is not implemented.");
            }

            throw new ArgumentException("Bad expression.");
        }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// Function call expression, e.g. sin(0.5)
    /// </summary>
    internal class FunctionExpression : Expression
    {
        public FunctionExpression(FunctionDefinition function, List<Expression> parms) : this(null, function, parms)
        { }

        public FunctionExpression(PrefixNode prefix, FunctionDefinition function, List<Expression> parms)
        {
            Prefix = prefix;
            FunctionDefinition = function;
            Parameters = parms;
        }

        // ** object model
        override public object Evaluate()
        {
            return FunctionDefinition.Function(Parameters);
        }

        public PrefixNode Prefix { get; }

        public FunctionDefinition FunctionDefinition { get; }

        public List<Expression> Parameters { get; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
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

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
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

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
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

        private readonly ExpressionErrorType _errorType;

        internal ErrorExpression(ExpressionErrorType errorType)
        {
            _errorType = errorType;
        }

        public override object Evaluate()
        {
            return _errorType;
        }

        public void ThrowApplicableException()
        {
            switch (_errorType)
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

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// An placeholder node for AST nodes that are not yet supported in ClosedXML.
    /// </summary>
    internal class NotSupportedNode : Expression
    {
        private readonly string _featureText;

        public NotSupportedNode(string featureText)
        {
            _featureText = featureText;
        }

        public override object Evaluate()
        {
            throw new NotImplementedException($"Evaluation of {_featureText} is not implemented.");
        }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// AST node for an reference to an external file in a formula.
    /// </summary>
    internal class FileNode : ExpressionBase
    {
        /// <summary>
        /// If the file is references indirectly, numeric identifier of a file.
        /// </summary>
        public int? Numeric { get; }

        /// <summary>
        /// If a file is referenced directly, a path to the file on the disc/UNC/web link, .
        /// </summary>
        public string Path { get; }

        public FileNode(string path)
        {
            Path = path;
        }

        public FileNode(int numeric)
        {
            Numeric = numeric;
        }

        public static void CreateFileNode(AstContext context, ParseTreeNode parseNode)
        {
            var filePath = string.Empty;
            FileNode fileNode = null;
            foreach (ParseTreeNode nt in parseNode.ChildNodes)
            {
                if (nt.Term.Name == GrammarNames.TokenFileNameNumeric)
                {
                    var numberInBrackets = nt.Token.ValueString;
                    var fileNumericIndex = int.Parse(numberInBrackets.Substring(1, numberInBrackets.Length - 2), NumberStyles.None);
                    fileNode = new FileNode(fileNumericIndex);
                    break;
                }

                switch (nt.Term.Name)
                {
                    case GrammarNames.TokenFilePath:
                        filePath = nt.Token.ValueString;
                        break;
                    case GrammarNames.TokenFileNameEnclosedInBrackets:
                        fileNode = new FileNode(System.IO.Path.Combine(filePath, nt.Token.ValueString));
                        break;
                    case GrammarNames.TokenFileName:
                        fileNode = new FileNode(System.IO.Path.Combine(filePath, nt.Token.ValueString));
                        break;
                    default:
                        throw new ArgumentOutOfRangeException($"Unexpected term {nt.Term.Name}.");
                }

            }
            parseNode.AstNode = fileNode;
        }
    }

    /// <summary>
    /// AST node for prefix of a reference in a formula. Prefix is a specification where to look for a reference.
    /// </summary>
    internal class PrefixNode : ExpressionBase
    {
        private PrefixNode(FileNode file, string sheet, string firstSheet, string lastSheet)
        {
            File = file;
            Sheet = sheet;
            FirstSheet = firstSheet;
            LastSheet = lastSheet;
        }

        /// <summary>
        /// If prefix references data from another file, can be empty.
        /// </summary>
        public FileNode File { get; }

        /// <summary>
        /// Name of the sheet, without ! or escaped quotes. Can be empty in some cases (e.g. reference to a named range in an another file).
        /// </summary>
        public string Sheet { get; }

        /// <summary>
        /// If the prefix is for 3D reference, name of first sheet. Empty otherwise.
        /// </summary>
        public string FirstSheet { get; }

        /// <summary>
        /// If the prefix is for 3D reference, name of the last sheet. Empty otherwise.
        /// </summary>
        public string LastSheet { get; }

        public static void CreatePrefixNode(AstContext context, ParseTreeNode parseNode)
        {
            PrefixNode prefix = null;
            FileNode fileNode = null;
            foreach (var nt in parseNode.ChildNodes)
            {
                if (nt.AstNode is FileNode fn)
                {
                    fileNode = fn;
                    continue;
                }

                switch (nt.Term.Name)
                {
                    case "'":
                        // Quoted sheet name has a single quote ' as first term.
                        continue;
                    case GrammarNames.TokenSheet:
                        var sheetName = RemoveExclamationMark(nt.Token.ValueString);
                        prefix = new PrefixNode(fileNode, sheetName, null, null);
                        break;

                    case GrammarNames.TokenSheetQuoted:
                        var quotedSheetName = RemoveExclamationMark("'" + nt.Token.ValueString);
                        prefix = new PrefixNode(fileNode, quotedSheetName.UnescapeSheetName(), null, null);
                        break;

                    case GrammarNames.TokenRefError:
                        // #REF! is a valid sheet name, Token.ValueString is lower case for some reason.
                        prefix = new PrefixNode(fileNode, RemoveExclamationMark(nt.Token.Text), null, null);
                        break;
                    case GrammarNames.TokenMultipleSheets:
                        var normalSheets = RemoveExclamationMark(nt.Token.Text).Split(':');
                        prefix = new PrefixNode(fileNode, null, normalSheets[0], normalSheets[1]);
                        break;
                    case GrammarNames.TokenMultipleSheetsQuoted:
                        var quotedSheets = RemoveExclamationMark(nt.Token.Text).Split(':');
                        prefix = new PrefixNode(fileNode, null, quotedSheets[0], quotedSheets[1]);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException($"Unexpected term {nt.Term.Name}.");
                }
            }

            if (prefix is null)
                prefix = new PrefixNode(fileNode, null, null, null);

            parseNode.AstNode = prefix;

            static string RemoveExclamationMark(string sheetName) => sheetName.Substring(0, sheetName.Length - 1);
        }
    }

    /// <summary>
    /// AST node for a reference of an area in some sheet.
    /// </summary>
    internal class ReferenceNode : Expression
    {
        public ReferenceNode(PrefixNode prefix, ReferenceItemType type, string address)
        {
            Prefix = prefix;
            Type = type;
            Address = address;
        }

        /// <summary>
        /// An optional prefix for reference item.
        /// </summary>
        public PrefixNode Prefix { get; }

        public ReferenceItemType Type { get; }

        /// <summary>
        /// An address of a reference that corresponds to <see cref="Type"/>.
        /// </summary>
        public string Address { get; }

        public override object Evaluate() => throw new NotImplementedException("Evaluation of reference is not implemented.");

        /// <summary>
        /// Reference AST node is significantly different from CST node. It takes Reference, ReferenceFunctionCall and ReferenceItem terms into a reference value
        /// that represent an area of a workbook (ReferenceNode, StructuredReferenceNode) and operations over these areas (BinaryOperation, UnaryOperation, FunctionExpression).
        /// </summary>
        public static void CreateReferenceNode(AstContext context, ParseTreeNode parseNode)
        {
            PrefixNode prefix = null;
            Expression referenceNode = null;
            foreach (var nt in parseNode.ChildNodes)
            {
                if (nt.AstNode is PrefixNode p)
                {
                    prefix = p;
                    continue;
                }

                // Copy node from ReferenceFunctionCall: 'Reference + colon + Reference', 'Reference + intersectop + Reference', 'OpenParen + Union + CloseParen', 'RefFunctionName + Arguments + CloseParen'
                if (nt.AstNode is BinaryExpression binOp)
                {
                    referenceNode = binOp;
                    break;
                }

                // Copy node from ReferenceFunctionCall: 'Reference + hash'
                if (nt.AstNode is UnaryExpression unaryOp)
                {
                    referenceNode = unaryOp;
                    break;
                }

                // Copy node from ReferenceFunctionCall: 'RefFunctionName + Arguments + CloseParen' (never has prefix)
                // Copy node from ReferenceItem: 'UDFunctionCall' (can have prefix)
                if (nt.AstNode is FunctionExpression fn)
                {
                    referenceNode = new FunctionExpression(prefix, fn.FunctionDefinition, fn.Parameters);
                    break;
                }

                // Copy node from ReferenceItem: 'StructuredReference'
                if (nt.AstNode is StructuredReferenceNode)
                {
                    // TODO: Copy structured reference once implemented
                    referenceNode = new StructuredReferenceNode(prefix);
                    break;
                }

                // Copy node from ReferenceItem: 'RefError'
                if (nt.AstNode is ErrorExpression errorNode)
                {
                    // Although the #REF! can have prefix, there is no difference
                    referenceNode = errorNode;
                    break;
                }

                // Copy node from Reference: 'OpenParen + Reference + PreferShiftHere() + CloseParen'
                if (nt.AstNode is ReferenceNode rn)
                {
                    referenceNode = rn;
                    break;
                }

                switch (nt.Term.Name)
                {
                    case GrammarNames.Cell:
                        referenceNode = new ReferenceNode(prefix, ReferenceItemType.Cell, nt.ChildNodes.Single().Token.Text);
                        break;
                    case GrammarNames.NamedRange:
                        // Named range can be NameToken or NamedRangeCombinationToken. The second one is there only to detect names like A1A1.
                        referenceNode = new ReferenceNode(prefix, ReferenceItemType.NamedRange, nt.ChildNodes.Single().Token.Text);
                        break;
                    case GrammarNames.HorizontalRange:
                        referenceNode = new ReferenceNode(prefix, ReferenceItemType.HRange, nt.ChildNodes.Single().Token.Text);
                        break;
                    case GrammarNames.VerticalRange:
                        referenceNode = new ReferenceNode(prefix, ReferenceItemType.VRange, nt.ChildNodes.Single().Token.Text);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException($"Unexpected term {nt.Term.Name}.");
                }
            }

            parseNode.AstNode = referenceNode;
        }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    internal enum ReferenceItemType { Cell, NamedRange, VRange, HRange }

    // TODO: The AST node doesn't have any stuff from StructuredReference term because structured reference is not yet suported and
    // the SR grammar has changed in not-yet-released (after 1.5.2) version of XLParser
    internal class StructuredReferenceNode : Expression
    {
        public StructuredReferenceNode(PrefixNode prefix)
        {
            Prefix = prefix;
        }

        /// <summary>
        /// Can be empty if no prefix available.
        /// </summary>
        public PrefixNode Prefix { get; }

        public override object Evaluate() => throw new NotImplementedException("Evaluation of structured references is not implemented.");

        public static void CreateStructuredReferenceNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new StructuredReferenceNode(null);
        }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
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
