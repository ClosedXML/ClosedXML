using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Range>;

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

        public AnyValue Visit(CalcContext context, ErrorExpression node)
        {
            throw new NotImplementedException();
        }

        public AnyValue Visit(CalcContext context, UnaryExpression node)
        {
            var arg = node.Expression.Accept(context, this);
            if (context.UseImplicitIntersection)
            {
                arg = arg.ImplicitIntersection(context);
            }

            return node.Operation switch
            {
                UnaryOp.Add => arg.UnaryPlus(),
                UnaryOp.Subtract => arg.UnaryMinus(context),
                UnaryOp.Percentage => arg.UnaryPercent(context),
                UnaryOp.SpillRange => throw new NotImplementedException(),
                UnaryOp.ImplicitIntersection => throw new NotImplementedException("Excel 2016 implicit intersection is different from @ intersection of E2019+"),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, BinaryExpression node)
        {
            var leftArg = node.LeftExpression.Accept(context, this);
            var rightArg = node.RightExpression.Accept(context, this);

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
                BinaryOp.Concat => throw new NotImplementedException(),
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
            var args = new AnyValue[node.Parameters.Count];
            for (var i = 0; i < node.Parameters.Count; ++i)
                args[i] = node.Parameters[i].Accept(context, this);

            if (!_functions.TryGetFunc(node.Name, out FormulaFunction function))
            {
                if (!_functions.TryGetFunc(node.Name, out FunctionDefinition legacyFunction))
                    return Error1.Name;

                var rangeFunctions = new Dictionary<string, List<int>>()
                {
                    { "AND" , new List<int>{ 0 } }
                };
                rangeFunctions.TryGetValue(node.Name, out var ignoreIdx);
                for (var i = 0; i < args.Length; ++i)
                {
                    if (ignoreIdx is not null && ignoreIdx.Contains(i))
                    {
                    }
                    else
                    {
                        args[i] = args[i].ImplicitIntersection(context);
                    }
                }

                // This creates a some of overhead, but all legacy functions will be migrated in near future
                var adaptedArgs = new List<Expression>(args.Length);
                foreach (var arg in args)
                {
                    var adaptedArg = arg.Match<Expression>(
                        logical => new AdapterExpression(logical.Value),
                        number => new AdapterExpression(number.Value),
                        text => new AdapterExpression(text.Value),
                        error => new ErrorExpression(error.Type),
                        array => throw new NotSupportedException("Legacy CalcEngine couldn't work with arrays and neither will adapter."),
                        range =>
                        {
                            if (range.Areas.Count != 1)
                                throw new NotSupportedException($"Legacy XObjectExpression could only work with single area, reference has {range.Areas.Count}.");

                            var area = range.Areas.Single();
                            return new XObjectExpression(new CellRangeReference(context.Worksheet.Range(area), context.CalcEngine));
                        });
                    adaptedArgs.Add(adaptedArg);
                }
                try
                {
                    var result = legacyFunction.Function(adaptedArgs);
                    return result switch
                    {
                        bool logic => AnyValue.FromT0(new Logical(logic)),
                        double number => AnyValue.FromT1(new Number1(number)),
                        string text => AnyValue.FromT2(new Text(text)),
                        _ => throw new NotImplementedException($"Got a result from some function type {result?.GetType().Name ?? "null"} with value {result}.")
                    };
                }
                catch (CalcEngineException)
                {
                    //TODO: Map errors
                    throw;
                }
            }

            return function.CallFunction(context, args);
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
                worksheet = context.Worksheet;
            }

            return new Range(new XLRangeAddress(worksheet, node.Address));
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

        public AnyValue Visit(CalcContext context, XObjectExpression node) => throw new InvalidOperationException();

        public AnyValue Visit(CalcContext context, EmptyValueExpression node) => throw new InvalidOperationException();

        #endregion

        private class AdapterExpression : Expression
        {
            private readonly object _value;

            public AdapterExpression(object value)
            {
                _value = value;
            }

            public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor)
                => throw new InvalidOperationException("The node should never be used in visitor.");

            public override object Evaluate() => _value;
        }
    }
}
