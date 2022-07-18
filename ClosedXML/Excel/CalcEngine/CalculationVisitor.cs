using System;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference1>;

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
            switch (node.Evaluate())
            {
                case double number:
                    return new Number1(number);
                case bool logical:
                    return new Logical(logical);
                case string text:
                    return new Text(text);
                default:
                    throw new InvalidOperationException();
            }
        }

        public AnyValue Visit(CalcContext context, UnaryExpression node)
        {
            var arg = node.Expression.Accept(context, this);
            return node.Operation switch
            {
                UnaryOp.Add => arg.UnaryPlus(),
                UnaryOp.Subtract => arg.UnaryMinus(context),
                UnaryOp.Percentage => arg.UnaryPercent(context),
                UnaryOp.SpillRange => throw new NotImplementedException(),
                UnaryOp.ImplicitIntersection => throw new NotImplementedException("E2016 implicit intersection is different from @ intersection of E2019+"), // arg.ImplicitIntersection(context),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, BinaryExpression node)
        {
            var leftArg = node.LeftExpression.Accept(context, this);
            var rightArg = node.RightExpression.Accept(context, this);

            // References operators
            switch (node.Operation)
            {
                case BinaryOp.Range: return leftArg.ReferenceRange(rightArg);
                case BinaryOp.Union: return leftArg.ReferenceUnion(rightArg);
                case BinaryOp.Intersection: throw new NotImplementedException();
            };

            if (context.UseImplicitIntersection)
            {
                leftArg = leftArg.ImplicitIntersection(context);
                rightArg = rightArg.ImplicitIntersection(context);
            }

            return node.Operation switch
            {
                // Text operators
                BinaryOp.Concat => throw new NotImplementedException(),
                // Arithmetic
                BinaryOp.Add => leftArg.BinaryPlus(rightArg, context),
                BinaryOp.Sub => leftArg.BinaryMinus(rightArg, context),
                BinaryOp.Mult => leftArg.BinaryMult(rightArg, context),
                BinaryOp.Div => leftArg.BinaryDiv(rightArg, context),
                BinaryOp.Exp => leftArg.BinaryExp(rightArg, context),
                // Comparison operators
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
            if (!_functions.TryGetFunc(node.Name, out var function))
                return Error1.Name;

            var args = new AnyValue[node.Parameters.Count];
            for (var i = 0; i < node.Parameters.Count; ++i)
            {
                args[i] = node.Parameters[i].Accept(context, this);
            }

            return function.CallFunction(context, args);
        }

        public AnyValue Visit(CalcContext context, XObjectExpression node)
        {
            throw new NotImplementedException();
        }

        public AnyValue Visit(CalcContext context, EmptyValueExpression node)
        {
            throw new NotImplementedException();
        }

        public AnyValue Visit(CalcContext context, ErrorExpression node)
        {
            throw new NotImplementedException();
        }

        public AnyValue Visit(CalcContext context, NotSupportedNode node)
        {
            throw new NotImplementedException();
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
                if (!context.Worksheet.Workbook.TryGetWorksheet(sheet, out var worksheet1))
                    return Error1.Ref;
                worksheet = (XLWorksheet)worksheet1;
            }
            else
            {
                // FIXME: This is bad, it won't work with other sheets
                worksheet = (XLWorksheet)context.Worksheet;
            }

            return new Reference1(new XLRangeAddress(worksheet, node.Address));
        }

        public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
        {
            throw new NotImplementedException();
        }

        public AnyValue Visit(CalcContext context, PrefixNode node)
        {
            throw new NotImplementedException();
        }

        public AnyValue Visit(CalcContext context, FileNode node)
        {
            throw new NotImplementedException();
        }
    }
}
