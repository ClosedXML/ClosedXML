using System;
using System.Collections.Generic;
using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FormulaParser
    {
        private readonly AstFactoryA1 _nodeFactory;

        public FormulaParser(FunctionRegistry functionRegistry)
        {
            _nodeFactory = new AstFactoryA1(functionRegistry);
        }

        /// <summary>
        /// Parse a formula into an abstract syntax tree.
        /// </summary>
        public Formula GetAst(string formula)
        {
            // Equality sign at the beginning of formula is only visualization in the GUI, real formulas don't have it.
            if (formula.Length > 0 && formula[0] == '=')
                formula = formula.Substring(1);

            try
            {
                var ctx = new List<FormulaFlags>();
                var root = FormulaParser<ScalarValue, ValueNode, List<FormulaFlags>>.CellFormulaA1(formula, ctx,
                    _nodeFactory);
                var flags = ctx.Contains(FormulaFlags.HasSubtotal)
                    ? FormulaFlags.HasSubtotal
                    : FormulaFlags.None;
                return new Formula(formula, root, flags);
            }
            catch (ParsingException ex)
            {
                throw new ExpressionParseException(ex.Message);
            }
        }

        /// <summary>
        /// Factory to create abstract syntax tree for a formula in A1 notation.
        /// </summary>
        private sealed class AstFactoryA1 : IAstFactory<ScalarValue, ValueNode, List<FormulaFlags>>
        {
            /// <summary>
            /// A prefix for so-called future functions. Excel can add functions, but to avoid name collisions,
            /// it prefixes names of function with this prefix. The prefix is omitted from GUI.
            /// </summary>
            /// <example>
            /// If you write <c>CONCAT(A1,B1)</c> in Excel 2021 (not present in Excel 2013), it is saved to the
            /// worksheet file as <c>_xlfn.CONCAT(A1,B1)</c>, but the Excel GUI will show only <c>CONCAT(A1,B1)</c>,
            /// without the <c>_xlfn</c>.
            /// </example>
            private const string DefaultFunctionNameSpace = "_xlfn";

            private readonly FunctionRegistry _functionRegistry;

            internal AstFactoryA1(FunctionRegistry functionRegistry)
            {
                _functionRegistry = functionRegistry;
            }

            public ScalarValue LogicalValue(List<FormulaFlags> context, bool logical) => logical;

            public ScalarValue NumberValue(List<FormulaFlags> context, double number) => number;

            public ScalarValue TextValue(List<FormulaFlags> context, string text) => text;

            public ScalarValue ErrorValue(List<FormulaFlags> context, ReadOnlySpan<char> errorText)
            {
                return GetErrorValue(errorText);
            }

            public ValueNode ArrayNode(List<FormulaFlags> context, int rows, int columns,
                IReadOnlyList<ScalarValue> elements)
            {
                var array = new LiteralArray(rows, columns, elements);
                return new ArrayNode(array);
            }

            public ValueNode BlankNode(List<FormulaFlags> context)
            {
                return new ScalarNode(ScalarValue.Blank);
            }

            public ValueNode LogicalNode(List<FormulaFlags> context, bool logical)
            {
                return new ScalarNode(logical);
            }

            public ValueNode ErrorNode(List<FormulaFlags> context, ReadOnlySpan<char> errorText)
            {
                var error = GetErrorValue(errorText);
                return new ScalarNode(error);
            }

            public ValueNode NumberNode(List<FormulaFlags> context, double number)
            {
                return new ScalarNode(number);
            }

            public ValueNode TextNode(List<FormulaFlags> context, string text)
            {
                return new ScalarNode(text);
            }

            public ValueNode Reference(List<FormulaFlags> context, ReferenceArea area)
            {
                return new ReferenceNode(null, area);
            }

            public ValueNode SheetReference(List<FormulaFlags> context, string sheet, ReferenceArea area)
            {
                var prefixNode = new PrefixNode(null, sheet, null, null);
                return new ReferenceNode(prefixNode, area);
            }

            public ValueNode Reference3D(List<FormulaFlags> context, string firstSheet, string lastSheet,
                ReferenceArea area)
            {
                var prefixNode = new PrefixNode(null, null, firstSheet, lastSheet);
                return new ReferenceNode(prefixNode, area);
            }

            public ValueNode ExternalSheetReference(List<FormulaFlags> context, int workbookIndex, string sheet,
                ReferenceArea area)
            {
                var fileNode = new FileNode(workbookIndex);
                var prefixNode = new PrefixNode(fileNode, sheet, null, null);
                return new ReferenceNode(prefixNode, area);
            }

            public ValueNode ExternalReference3D(List<FormulaFlags> context, int workbookIndex, string firstSheet,
                string lastSheet, ReferenceArea area)
            {
                var fileNode = new FileNode(workbookIndex);
                var prefixNode = new PrefixNode(fileNode, null, firstSheet, lastSheet);
                return new ReferenceNode(prefixNode, area);
            }

            public ValueNode Function(List<FormulaFlags> context, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var functionName = name.ToString();
                return GetFunctionNode(context, null, functionName, args);
            }

            public ValueNode Function(List<FormulaFlags> context, string sheetName, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(null, sheetName, null, null);
                return GetFunctionNode(context, prefixNode, name.ToString(), args);
            }

            public ValueNode ExternalFunction(List<FormulaFlags> context, int workbookIndex, string sheet,
                ReadOnlySpan<char> name, IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet, null, null);
                return GetFunctionNode(context, prefixNode, name.ToString(), args);
            }

            public ValueNode ExternalFunction(List<FormulaFlags> context, int workbookIndex, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
                return GetFunctionNode(context, prefixNode, name.ToString(), args);
            }

            public ValueNode CellFunction(List<FormulaFlags> context, Parser.Reference cell,
                IReadOnlyList<ValueNode> args)
            {
                return new NotSupportedNode("Cell functions are not yet supported.");
            }

            public ValueNode StructureReference(List<FormulaFlags> context, StructuredReferenceArea area,
                string? firstColumn, string? lastColumn)
            {
                return new StructuredReferenceNode(null, null, area, firstColumn, lastColumn);
            }

            public ValueNode StructureReference(List<FormulaFlags> context, string table, StructuredReferenceArea area,
                string? firstColumn, string? lastColumn)
            {
                return new StructuredReferenceNode(null, table, area, firstColumn, lastColumn);
            }

            public ValueNode ExternalStructureReference(List<FormulaFlags> context, int workbookIndex, string table,
                StructuredReferenceArea area, string? firstColumn, string? lastColumn)
            {
                return new StructuredReferenceNode(new PrefixNode(new FileNode(workbookIndex), null, null, null), table,
                    area, firstColumn, lastColumn);
            }

            public ValueNode Name(List<FormulaFlags> context, string name)
            {
                return new NameNode(null, name);
            }

            public ValueNode SheetName(List<FormulaFlags> context, string sheet, string name)
            {
                var prefixNode = new PrefixNode(null, sheet, null, null);
                return new NameNode(prefixNode, name);
            }

            public ValueNode ExternalName(List<FormulaFlags> context, int workbookIndex, string name)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
                return new NameNode(prefixNode, name);
            }

            public ValueNode ExternalSheetName(List<FormulaFlags> context, int workbookIndex, string sheet, string name)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet, null, null);
                return new NameNode(prefixNode, name);
            }

            public ValueNode BinaryNode(List<FormulaFlags> context, BinaryOperation operation, ValueNode leftNode,
                ValueNode rightNode)
            {
                var op = operation switch
                {
                    BinaryOperation.Concat => BinaryOp.Concat,
                    BinaryOperation.GreaterOrEqualThan => BinaryOp.Gte,
                    BinaryOperation.LessOrEqualThan => BinaryOp.Lte,
                    BinaryOperation.LessThan => BinaryOp.Lt,
                    BinaryOperation.GreaterThan => BinaryOp.Gt,
                    BinaryOperation.NotEqual => BinaryOp.Neq,
                    BinaryOperation.Equal => BinaryOp.Eq,
                    BinaryOperation.Addition => BinaryOp.Add,
                    BinaryOperation.Subtraction => BinaryOp.Sub,
                    BinaryOperation.Multiplication => BinaryOp.Mult,
                    BinaryOperation.Division => BinaryOp.Div,
                    BinaryOperation.Power => BinaryOp.Exp,
                    BinaryOperation.Union => BinaryOp.Union,
                    BinaryOperation.Intersection => BinaryOp.Intersection,
                    BinaryOperation.Range => BinaryOp.Range,
                    _ => throw new NotSupportedException($"'{operation}' is not a binary operation.")
                };

                return new BinaryNode(op, leftNode, rightNode);
            }

            public ValueNode Unary(List<FormulaFlags> context, UnaryOperation operation, ValueNode node)
            {
                var op = operation switch
                {
                    UnaryOperation.Plus => UnaryOp.Add,
                    UnaryOperation.Minus => UnaryOp.Subtract,
                    UnaryOperation.Percent => UnaryOp.Percentage,
                    UnaryOperation.ImplicitIntersection => UnaryOp.ImplicitIntersection,
                    UnaryOperation.SpillRange => UnaryOp.SpillRange,
                    _ => throw new NotSupportedException($"'{operation}' is not a unary operation.")
                };
                return new UnaryNode(op, node);
            }

            public ValueNode Nested(List<FormulaFlags> context, ValueNode node)
            {
                return node;
            }

            private FunctionNode GetFunctionNode(List<FormulaFlags> context, PrefixNode? prefixNode, string functionName,
                IReadOnlyList<ValueNode> argumentNodes)
            {
                var foundFunction = _functionRegistry.TryGetFunc(functionName, out var minParams, out var maxParams);

                // If function is a future function, strip the prefix because all registration of functions
                // are without a prefix. That should change, but it's a reality for now.
                if (!foundFunction && functionName.StartsWith($"{DefaultFunctionNameSpace}."))
                {
                    functionName = functionName.Substring(DefaultFunctionNameSpace.Length + 1);
                    foundFunction = _functionRegistry.TryGetFunc(functionName, out minParams, out maxParams);
                }

                if (string.Equals(functionName, @"SUBTOTAL", StringComparison.OrdinalIgnoreCase))
                    context.Add(FormulaFlags.HasSubtotal);

                // Even if we haven't found anything, don't crash. Missing function will be evaluated to `#NAME?`
                if (!foundFunction)
                    return new FunctionNode(functionName, argumentNodes);

                if (minParams != -1 && argumentNodes.Count < minParams)
                    throw new ExpressionParseException(
                        $"Too few parameters for function '{functionName}'. Expected a minimum of {minParams} and a maximum of {maxParams}.");

                if (maxParams != -1 && argumentNodes.Count > maxParams)
                    throw new ExpressionParseException(
                        $"Too many parameters for function '{functionName}'.Expected a minimum of {minParams} and a maximum of {maxParams}.");

                return new FunctionNode(prefixNode, functionName, argumentNodes);
            }

            private static XLError GetErrorValue(ReadOnlySpan<char> error)
            {
                if (!XLErrorParser.TryParseError(error.ToString(), out var errorEnum))
                    throw new InvalidOperationException($"'{error.ToString()}' is not error.");
                return errorEnum;
            }
        }
    }
}
