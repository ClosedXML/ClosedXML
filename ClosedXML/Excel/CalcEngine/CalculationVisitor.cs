using System;
using System.Linq;
using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

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

        public AnyValue Visit(CalcContext context, UnaryExpression node)
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

        public AnyValue Visit(CalcContext context, BinaryExpression node)
        {
            var leftArg = node.LeftExpression.Accept(context, this);
            var rightArg = node.RightExpression.Accept(context, this);

            return node.Operation switch
            {
                BinaryOp.Range => leftArg.ReferenceRange(rightArg, context),
                BinaryOp.Union => leftArg.ReferenceUnion(rightArg),
                BinaryOp.Intersection => throw new NotImplementedException("Evaluation of range intersection operator is not implemented."),
                BinaryOp.Concat => leftArg.Concat(rightArg, context),
                BinaryOp.Add => leftArg.BinaryPlus(rightArg, context),
                BinaryOp.Sub => leftArg.BinaryMinus(rightArg, context),
                BinaryOp.Mult => leftArg.BinaryMult(rightArg, context),
                BinaryOp.Div => leftArg.BinaryDiv(rightArg, context),
                BinaryOp.Exp => leftArg.BinaryExp(rightArg, context),
                BinaryOp.Lt => leftArg.IsLessThan(rightArg, context),
                BinaryOp.Lte => leftArg.IsLessThanOrEqual(rightArg, context),
                BinaryOp.Eq => leftArg.IsEqual(rightArg, context),
                BinaryOp.Neq => leftArg.IsNotEqual(rightArg, context),
                BinaryOp.Gte => leftArg.IsGreaterThanOrEqual(rightArg, context),
                BinaryOp.Gt => leftArg.IsGreaterThan(rightArg, context),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, FunctionExpression node)
        {
            if (!_functions.TryGetFunc(node.Name, out FunctionDefinition fn))
                return Error.NameNotRecognized;

            var args = GetArgs(context, fn, node);
            return fn.CallFunction(context, args);
        }

        private AnyValue?[] GetArgs(CalcContext context, FunctionDefinition fn, FunctionExpression node)
        {
            var args = new AnyValue?[node.Parameters.Count];
            for (var argIndex = 0; argIndex < node.Parameters.Count; ++argIndex)
            {
                var paramNode = node.Parameters[argIndex];
                var arg = paramNode is not EmptyArgumentNode ? node.Parameters[argIndex].Accept(context, this) : default(AnyValue?);

                if (context.UseImplicitIntersection && fn.AllowRanges != AllowRange.All && arg.HasValue)
                {
                    switch (fn.AllowRanges)
                    {
                        case AllowRange.None:
                            arg = arg.Value.ImplicitIntersection(context);
                            break;
                        case AllowRange.Except:
                            if (fn.MarkedParams.Contains(argIndex))
                                arg = arg.Value.ImplicitIntersection(context);

                            break;
                        case AllowRange.Only:
                            if (!fn.MarkedParams.Contains(argIndex))
                                arg = arg.Value.ImplicitIntersection(context);

                            break;
                        default:
                            throw new InvalidOperationException();
                    }
                }

                args[argIndex] = arg;
            }

            return args;
        }

        public AnyValue Visit(CalcContext context, ReferenceNode node)
        {
            XLWorksheet worksheet;
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
            else
            {
                worksheet = null;
            }

            if (node.Type == ReferenceItemType.Cell || node.Type == ReferenceItemType.HRange || node.Type == ReferenceItemType.VRange)
                return new Reference(new XLRangeAddress(worksheet, node.Address));

            var rangeName = node.Address;
            worksheet ??= context.Worksheet;
            if (!TryGetNamedRange(worksheet, rangeName, out var namedRange))
                return Error.NameNotRecognized;

            // This is rather horrible, but basically copy from XLCalcEngine.GetExternalObject
            // It's hard to count all things that are wrong with this, from hand parsing operator range union by XLNamedRange to recursion.
            if (!namedRange.IsValid)
                return Error.CellReference;

            // union is one of nodes that can't be in the root. Enclose in braces to make parser happy
            // TODO: Shoudl it always start with equal or never?
            var namedRangeFormula = namedRange.ToString();
            namedRangeFormula = !namedRangeFormula.StartsWith("=") ? "=(" + namedRange.ToString() + ")" : namedRangeFormula;
            var rangeResult = context.CalcEngine.EvaluateExpression(namedRangeFormula, context.Workbook, context.Worksheet);
            return rangeResult;

            static bool TryGetNamedRange(IXLWorksheet ws, string name, out XLNamedRange range)
            {
                var found = ws.NamedRanges.TryGetValue(name, out var namedRange)
                                    || ws.Workbook.NamedRanges.TryGetValue(name, out namedRange);
                range = (XLNamedRange)namedRange;
                return found;
            };
        }

        public AnyValue Visit(CalcContext context, NotSupportedNode node)
        {
            throw new NotImplementedException($"Evaluation of {node.FeatureName} is not implemented.");
        }

        public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
        {
            throw new NotImplementedException($"Evaluation of structured references is not implemented.");
        }

        #region Never visited nodes

        public AnyValue Visit(CalcContext context, PrefixNode node) => throw new InvalidOperationException();

        public AnyValue Visit(CalcContext context, FileNode node) => throw new NotImplementedException();

        public AnyValue Visit(CalcContext context, EmptyArgumentNode node) => throw new InvalidOperationException();

        #endregion
    }
}
