using System;
using System.Collections.Generic;
using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FormulaParser
    {
        private readonly AstFactory _nodeFactoryA1;
        private readonly AstFactory _nodeFactoryR1C1;

        public FormulaParser(FunctionRegistry functionRegistry)
        {
            _nodeFactoryA1 = new AstFactory(functionRegistry, true);
            _nodeFactoryR1C1 = new AstFactory(functionRegistry, false);
        }

        /// <summary>
        /// Parse a formula into an abstract syntax tree.
        /// </summary>
        public Formula GetAst(string formula, bool isA1)
        {
            // Equality sign at the beginning of formula is only visualization in the GUI, real formulas don't have it.
            if (formula.Length > 0 && formula[0] == '=')
                formula = formula.Substring(1);

            try
            {
                var ctx = new List<FormulaFlags>();
                
                var root = isA1
                    ? FormulaParser<ScalarValue, ValueNode, List<FormulaFlags>>.CellFormulaA1(formula, ctx, _nodeFactoryA1)
                    : FormulaParser<ScalarValue, ValueNode, List<FormulaFlags>>.CellFormulaR1C1(formula, ctx, _nodeFactoryR1C1);
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
        private sealed class AstFactory : IAstFactory<ScalarValue, ValueNode, List<FormulaFlags>>
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
            private readonly bool _isA1;

            internal AstFactory(FunctionRegistry functionRegistry, bool isA1)
            {
                _functionRegistry = functionRegistry;
                _isA1 = isA1;
            }

            public ScalarValue LogicalValue(List<FormulaFlags> context, SymbolRange range, bool logical) => logical;

            public ScalarValue NumberValue(List<FormulaFlags> context, SymbolRange range, double number) => number;

            public ScalarValue TextValue(List<FormulaFlags> context, SymbolRange range, string text) => text;

            public ScalarValue ErrorValue(List<FormulaFlags> context, SymbolRange range, ReadOnlySpan<char> errorText)
            {
                return GetErrorValue(errorText);
            }

            public ValueNode ArrayNode(List<FormulaFlags> context, SymbolRange range, int rows, int columns,
                IReadOnlyList<ScalarValue> elements)
            {
                var array = new LiteralArray(rows, columns, elements);
                return new ArrayNode(array);
            }

            public ValueNode BlankNode(List<FormulaFlags> context, SymbolRange range)
            {
                return new ScalarNode(ScalarValue.Blank);
            }

            public ValueNode LogicalNode(List<FormulaFlags> context, SymbolRange range, bool logical)
            {
                return new ScalarNode(logical);
            }

            public ValueNode ErrorNode(List<FormulaFlags> context, SymbolRange range, ReadOnlySpan<char> errorText)
            {
                var error = GetErrorValue(errorText);
                return new ScalarNode(error);
            }

            public ValueNode NumberNode(List<FormulaFlags> context, SymbolRange range, double number)
            {
                return new ScalarNode(number);
            }

            public ValueNode TextNode(List<FormulaFlags> context, SymbolRange range, string text)
            {
                return new ScalarNode(text);
            }

            public ValueNode Reference(List<FormulaFlags> context, SymbolRange range, ReferenceArea area)
            {
                return new ReferenceNode(null, area, _isA1);
            }

            public ValueNode SheetReference(List<FormulaFlags> context, SymbolRange range, string sheet, ReferenceArea area)
            {
                var prefixNode = new PrefixNode(null, sheet, null, null);
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode Reference3D(List<FormulaFlags> context, SymbolRange range, string firstSheet, string lastSheet,
                ReferenceArea area)
            {
                var prefixNode = new PrefixNode(null, null, firstSheet, lastSheet);
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode ExternalSheetReference(List<FormulaFlags> context, SymbolRange range, int workbookIndex, string sheet,
                ReferenceArea area)
            {
                var fileNode = new FileNode(workbookIndex);
                var prefixNode = new PrefixNode(fileNode, sheet, null, null);
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode ExternalReference3D(List<FormulaFlags> context, SymbolRange range, int workbookIndex, string firstSheet,
                string lastSheet, ReferenceArea area)
            {
                var fileNode = new FileNode(workbookIndex);
                var prefixNode = new PrefixNode(fileNode, null, firstSheet, lastSheet);
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode Function(List<FormulaFlags> context, SymbolRange range, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var functionName = name.ToString();
                return GetFunctionNode(context, null, functionName, args);
            }

            public ValueNode Function(List<FormulaFlags> context, SymbolRange range, string sheetName, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(null, sheetName, null, null);
                return GetFunctionNode(context, prefixNode, name.ToString(), args);
            }

            public ValueNode ExternalFunction(List<FormulaFlags> context, SymbolRange range, int workbookIndex, string sheet,
                ReadOnlySpan<char> name, IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet, null, null);
                return GetFunctionNode(context, prefixNode, name.ToString(), args);
            }

            public ValueNode ExternalFunction(List<FormulaFlags> context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
                return GetFunctionNode(context, prefixNode, name.ToString(), args);
            }

            public ValueNode CellFunction(List<FormulaFlags> context, SymbolRange range, RowCol cell,
                IReadOnlyList<ValueNode> args)
            {
                return new NotSupportedNode("Cell functions are not yet supported.");
            }

            public ValueNode StructureReference(List<FormulaFlags> context, SymbolRange range, StructuredReferenceArea area,
                string? firstColumn, string? lastColumn)
            {
                return new StructuredReferenceNode(null, null, area, firstColumn, lastColumn);
            }

            public ValueNode StructureReference(List<FormulaFlags> context, SymbolRange range, string table, StructuredReferenceArea area,
                string? firstColumn, string? lastColumn)
            {
                return new StructuredReferenceNode(null, table, area, firstColumn, lastColumn);
            }

            public ValueNode ExternalStructureReference(List<FormulaFlags> context, SymbolRange range, int workbookIndex, string table,
                StructuredReferenceArea area, string? firstColumn, string? lastColumn)
            {
                return new StructuredReferenceNode(new PrefixNode(new FileNode(workbookIndex), null, null, null), table,
                    area, firstColumn, lastColumn);
            }

            public ValueNode Name(List<FormulaFlags> context, SymbolRange range, string name)
            {
                return new NameNode(null, name);
            }

            public ValueNode SheetName(List<FormulaFlags> context, SymbolRange range, string sheet, string name)
            {
                var prefixNode = new PrefixNode(null, sheet, null, null);
                return new NameNode(prefixNode, name);
            }

            public ValueNode ExternalName(List<FormulaFlags> context, SymbolRange range, int workbookIndex, string name)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
                return new NameNode(prefixNode, name);
            }

            public ValueNode ExternalSheetName(List<FormulaFlags> context, SymbolRange range, int workbookIndex, string sheet, string name)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet, null, null);
                return new NameNode(prefixNode, name);
            }

            public ValueNode BinaryNode(List<FormulaFlags> context, SymbolRange range, BinaryOperation operation, ValueNode leftNode,
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

            public ValueNode Unary(List<FormulaFlags> context, SymbolRange range, UnaryOperation operation, ValueNode node)
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

            public ValueNode Nested(List<FormulaFlags> context, SymbolRange range, ValueNode node)
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
