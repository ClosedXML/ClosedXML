using System;
using System.Buffers;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
    {
        private readonly FunctionRegistry _functions;
        private readonly ArrayPool<AnyValue> _argsPool;
        public CalculationVisitor(FunctionRegistry functions)
        {
            _functions = functions;
            _argsPool = ArrayPool<AnyValue>.Create(XLConstants.MaxFunctionArguments, 100);
        }

        public AnyValue Visit(CalcContext context, ScalarNode node)
        {
            return node.Value;
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

        public AnyValue Visit(CalcContext context, FunctionNode functionNode)
        {
            if (!_functions.TryGetFunc(functionNode.Name, out var fn))
                return XLError.NameNotRecognized;

            var parameters = functionNode.Parameters;
            var pool = _argsPool.Rent(parameters.Count);
            var args = new Span<AnyValue>(pool, 0, parameters.Count);
            try
            {
                for (var i = 0; i < parameters.Count; ++i)
                    args[i] = parameters[i].Accept(context, this);

                return fn.CallFunction(context, args);
            }
            finally
            {
                _argsPool.Return(pool);
            }
        }

        public AnyValue Visit(CalcContext context, ReferenceNode node)
        {
            return node.GetReference(context);
        }

        public AnyValue Visit(CalcContext context, NameNode node)
        {
            return node.GetValue(context.Worksheet, context.CalcEngine);
        }

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
