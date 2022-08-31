using System;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
    {
        public AnyValue Visit(CalcContext context, ScalarNode node)
        {
            // TODO: Refactor ScalarNode to a typed value instead of object value.
            return node.Evaluate() switch
            {
                bool logical => logical,
                int number => number,
                double number => number,
                string text => text,
                Error error => error,
                _ => throw new InvalidOperationException("Not a scalar value type")
            };
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
                BinaryOp.Range => throw new NotImplementedException("Evaluation of range operator is not implemented."),
                BinaryOp.Union => throw new NotImplementedException("Evaluation of range union operator is not implemented."),
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

        public AnyValue Visit(CalcContext context, XObjectExpression node)
        {
            throw new NotImplementedException($"Evaluation of a reference is not implemented.");
        }

        public AnyValue Visit(CalcContext context, FunctionNode node)
        {
            throw new NotImplementedException($"Evaluation of a reference is not implemented.");
        }

        public AnyValue Visit(CalcContext context, ReferenceNode node)
        {
            throw new NotImplementedException($"Evaluation of a reference is not implemented.");
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

        public AnyValue Visit(CalcContext context, FileNode node) => throw new InvalidOperationException();

        public AnyValue Visit(CalcContext context, EmptyArgumentNode node) => throw new InvalidOperationException();

        #endregion
    }
}
