using System;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
    {
        private readonly FunctionRegistry _functions;

        public CalculationVisitor(FunctionRegistry functions)
        {
            _functions = functions;
        }

        public AnyValue Visit(CalcContext context, ScalarNode node)
        {
            return node.Value;
        }

        public AnyValue Visit(CalcContext context, ErrorNode node)
        {
            return node.Error;
        }

        public AnyValue Visit(CalcContext context, UnaryNode node)
        {
            var arg = node.Expression.Accept(context, this);

            return node.Operation switch
            {
                UnaryOp.Add => arg.UnaryPlus(),
                UnaryOp.Subtract => arg.UnaryMinus(context),
                UnaryOp.Percentage => arg.UnaryPercent(context),
                UnaryOp.SpillRange => throw new NotImplementedException("Evaluation of spill range operator is not implemented."),
                UnaryOp.ImplicitIntersection => throw new NotImplementedException("Excel 2016 implicit intersection is different from @ intersection of E2019+"),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, BinaryNode node)
        {
            var leftArg = node.LeftExpression.Accept(context, this);
            var rightArg = node.RightExpression.Accept(context, this);

            return node.Operation switch
            {
                BinaryOp.Range => AnyValue.ReferenceRange(leftArg, rightArg, context),
                BinaryOp.Union => AnyValue.ReferenceUnion(leftArg, rightArg),
                BinaryOp.Intersection => throw new NotImplementedException("Evaluation of range intersection operator is not implemented."),
                BinaryOp.Concat => AnyValue.Concat(leftArg, rightArg, context),
                BinaryOp.Add => AnyValue.BinaryPlus(leftArg, rightArg, context),
                BinaryOp.Sub => AnyValue.BinaryMinus(leftArg, rightArg, context),
                BinaryOp.Mult => AnyValue.BinaryMult(leftArg, rightArg, context),
                BinaryOp.Div => AnyValue.BinaryDiv(leftArg, rightArg, context),
                BinaryOp.Exp => AnyValue.BinaryExp(leftArg, rightArg, context),
                BinaryOp.Lt => AnyValue.IsLessThan(leftArg, rightArg, context),
                BinaryOp.Lte => AnyValue.IsLessThanOrEqual(leftArg, rightArg, context),
                BinaryOp.Eq => AnyValue.IsEqual(leftArg, rightArg, context),
                BinaryOp.Neq => AnyValue.IsNotEqual(leftArg, rightArg, context),
                BinaryOp.Gte => AnyValue.IsGreaterThanOrEqual(leftArg, rightArg, context),
                BinaryOp.Gt => AnyValue.IsGreaterThan(leftArg, rightArg, context),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, FunctionNode node)
        {
            if (!_functions.TryGetFunc(node.Name, out FunctionDefinition fn))
                return Error.NameNotRecognized;

            var args = GetArgs(context, node);
            return fn.CallFunction(context, args);
        }

        private AnyValue[] GetArgs(CalcContext context, FunctionNode node)
        {
            var args = new AnyValue[node.Parameters.Count];
            for (var argIndex = 0; argIndex < node.Parameters.Count; ++argIndex)
            {
                var arg = node.Parameters[argIndex].Accept(context, this);
                args[argIndex] = arg;
            }

            return args;
        }

        public AnyValue Visit(CalcContext context, ReferenceNode node)
        {
            XLWorksheet worksheet = null;
            if (node.Prefix is not null)
            {
                if (node.Prefix.File is not null)
                    throw new NotImplementedException("References from other files are not yet implemented.");

                if (node.Prefix.FirstSheet is not null || node.Prefix.LastSheet is not null)
                    throw new NotImplementedException("3D references are not yet implemented.");

                var sheet = node.Prefix.Sheet;
                if (!context.Workbook.TryGetWorksheet(sheet, out var worksheet1))
                    return Error.CellReference;
                worksheet = (XLWorksheet)worksheet1;
            }

            if (node.Type == ReferenceItemType.Cell || node.Type == ReferenceItemType.HRange || node.Type == ReferenceItemType.VRange)
                return new Reference(new XLRangeAddress(worksheet, node.Address));

            // Only reference of type range left
            var rangeName = node.Address;
            worksheet ??= context.Worksheet;
            if (!TryGetNamedRange(worksheet, rangeName, out var namedRange))
                return Error.NameNotRecognized;

            // This is rather horrible, but basically copy from XLCalcEngine.GetExternalObject
            // It's hard to count all things that are wrong with this, from hand parsing operator range union by XLNamedRange to recursion.
            if (!namedRange.IsValid)
                return Error.CellReference;

            // Union (can easily be in the range) is one of the nodes that can't be in the root. Enclose in braces to make parser happy
            // Range can be something like 1+2, not just a reference to some area.
            var namedRangeFormula = namedRange.ToString();
            namedRangeFormula = !namedRangeFormula.StartsWith("=") ? "=(" + namedRange + ")" : namedRangeFormula;
            var rangeResult = context.CalcEngine.EvaluateExpression(namedRangeFormula, context.Workbook, context.Worksheet);
            return rangeResult;

            static bool TryGetNamedRange(IXLWorksheet ws, string name, out XLNamedRange range)
            {
                var found = ws.NamedRanges.TryGetValue(name, out var namedRange)
                                    || ws.Workbook.NamedRanges.TryGetValue(name, out namedRange);
                range = (XLNamedRange)namedRange;
                return found;
            }
        }

        public AnyValue Visit(CalcContext context, EmptyArgumentNode node)
            => AnyValue.Blank;

        public AnyValue Visit(CalcContext context, NotSupportedNode node)
            => throw new NotImplementedException($"Evaluation of {node.FeatureName} is not implemented.");

        public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
            => throw new NotImplementedException($"Evaluation of structured references is not implemented.");

        public AnyValue Visit(CalcContext context, PrefixNode node)
            => throw new InvalidOperationException("Node should never be visited.");

        public AnyValue Visit(CalcContext context, FileNode node)
            => throw new InvalidOperationException("Node should never be visited.");
    }
}
