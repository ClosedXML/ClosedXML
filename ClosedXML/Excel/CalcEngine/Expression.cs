using System;
using System.Collections;
using System.Collections.Generic;
using AnyValue = OneOf.OneOf<bool, ClosedXML.Excel.CalcEngine.Number1, string, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Base class for all AST nodes. All AST nodes must be immutable.
    /// </summary>
    internal abstract class AstNode
    {
        /// <summary>
        /// Method to accept a vistor (=call a method of visitor with correct type of the node).
        /// </summary>
        public abstract TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor);
    }

    /// <summary>
    /// A base class for all AST nodes that can be evaluated to produce a value.
    /// </summary>
    internal abstract class ValueNode : AstNode
    {
    }

    /// <summary>
    /// AST node that contains a number, text or a bool.
    /// </summary>
    internal class ScalarNode : ValueNode
    {
        public ScalarNode(AnyValue value)
        {
            Value = value;
        }

        public AnyValue Value { get; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    internal enum UnaryOp
    {
        Add,
        Subtract,
        Percentage,
        SpillRange,
        ImplicitIntersection
    }

    /// <summary>
    /// Unary expression, e.g. +123
    /// </summary>
    internal class UnaryExpression : ValueNode
    {
        public UnaryExpression(UnaryOp operation, ValueNode expr)
        {
            Operation = operation;
            Expression = expr;
        }

        public UnaryOp Operation { get; }

        public ValueNode Expression { get; private set; }

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
    internal class BinaryExpression : ValueNode
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

        public BinaryExpression(BinaryOp operation, ValueNode exprLeft, ValueNode exprRight)
        {
            _isComparison = _comparisons.Contains(operation);
            Operation = operation;
            LeftExpression = exprLeft;
            RightExpression = exprRight;
        }

        public BinaryOp Operation { get; }

        public ValueNode LeftExpression { get; private set; }

        public ValueNode RightExpression { get; private set; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// Function call expression, e.g. sin(0.5)
    /// </summary>
    internal class FunctionExpression : ValueNode
    {
        // TODO: Improve parser, node should have store a name, not a delegate
        public FunctionExpression(string name, List<ValueNode> parms) : this(null, name, parms)
        {
        }

        public FunctionExpression(PrefixNode prefix, string name, List<ValueNode> parms)
        {
            Prefix = prefix;
            Name = name;
            Parameters = parms;
        }

        public PrefixNode Prefix { get; }

        /// <summary>
        /// Name of the function.
        /// </summary>
        public string Name { get; }

        public List<ValueNode> Parameters { get; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// Expression that represents an external object.
    /// </summary>
    internal class XObjectExpression : Expression, IEnumerable
    {
        private readonly object _value;

        // ** ctor
        internal XObjectExpression(object value) : base(value)
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
    }

    /// <summary>
    /// Expression that represents an omitted parameter.
    /// </summary>
    internal class EmptyArgumentNode : ValueNode
    {
        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

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

    internal class ErrorExpression : ValueNode
    {
        internal ErrorExpression(ExpressionErrorType errorType)
        {
            ErrorType = errorType;
        }

        public ExpressionErrorType ErrorType { get; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// An placeholder node for AST nodes that are not yet supported in ClosedXML.
    /// </summary>
    internal class NotSupportedNode : ValueNode
    {
        public NotSupportedNode(string featureName)
        {
            FeatureName = featureName;
        }

        public string FeatureName { get; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// AST node for an reference to an external file in a formula.
    /// </summary>
    internal class FileNode : AstNode
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

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// AST node for prefix of a reference in a formula. Prefix is a specification where to look for a reference.
    /// <list type="bullet">
    /// <item>Prefix specifies a <c>Sheet</c> - used for references in the local workbook.</item>
    /// <item>Prefix specifies a <c>FirstSheet</c> and a <c>LastSheet</c> - 3D reference, references uses all sheets between first and last.</item>
    /// <item>Prefix specifies a <c>File</c>, no sheet is specified - used for named ranges in external file.</item>
    /// <item>Prefix specifies a <c>File</c> and a <c>Sheet</c> - references looks for its address in the sheet of the file.</item>
    /// </list>
    /// </summary>
    internal class PrefixNode : AstNode
    {
        public PrefixNode(FileNode file, string sheet, string firstSheet, string lastSheet)
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

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    /// <summary>
    /// AST node for a reference of an area in some sheet.
    /// </summary>
    internal class ReferenceNode : ValueNode
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
        /// An address of a reference that corresponds to <see cref="Type"/> or a name of named range.
        /// </summary>
        public string Address { get; }

        public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
    }

    internal enum ReferenceItemType { Cell, NamedRange, VRange, HRange }

    // TODO: The AST node doesn't have any stuff from StructuredReference term because structured reference is not yet suported and
    // the SR grammar has changed in not-yet-released (after 1.5.2) version of XLParser
    internal class StructuredReferenceNode : ValueNode
    {
        public StructuredReferenceNode(PrefixNode prefix)
        {
            Prefix = prefix;
        }

        /// <summary>
        /// Can be empty if no prefix available.
        /// </summary>
        public PrefixNode Prefix { get; }

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
