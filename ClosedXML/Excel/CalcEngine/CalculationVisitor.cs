using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
    {
        private readonly FunctionRegistry _functions;

        // WIP
        /// <summary>
        /// A (partial) list of functiosn that can accept arguments without implicit intersection.
        /// Key is the name of a function, the value is an 0-based index of function arguments that can accept range without implicit intersection. Null value means that all arguments can be range.
        /// </summary>
        private static readonly Dictionary<string, IReadOnlyList<int>> rangeFunctions = new(StringComparer.OrdinalIgnoreCase)
        {
            { "AND" , null },
            { "NETWORKDAYS", new List<int> { 2 } },
            { "WORKDAY", new List<int> { 2 } },
            { "AVERAGE", null },
            { "AVERAGEA", null },
            { "CONCAT", null },
            { "CONCATENATE", null }, // LEGACY: Remove after switch to new engine. This function actually doesn't acccept ranges, but it's legacy implementation has a check and there is a test.
            { "COUNT", null },
            { "COUNTA", null },
            { "COUNTBLANK", null},
            { "COUNTIF", new List<int>{ 0 } },
            { "COUNTIFs", Enumerable.Range(0, 255).Where(x => x % 2 == 0).ToList() },
            { "DEVSQ", null },
            { "GEOMEAN", null },
            { "HLOOKUP", new List<int>{ 1 } },
            { "INDEX", new List<int> { 0, 1 } },
            { "MATCH", new List<int> { 1 } },
            { "MAX", null },
            { "MAXA", null },
            { "MDETERM", new List<int>{ 0 } },
            { "MEDIAN", null },
            { "MIN", null },
            { "MINA", null },
            { "MINVERSE", null },
            { "MMULT", null },
            { "SERIESSUM", new List<int>{ 3 } }, // Yay, this function is part of ECMA-376, but isn't in the list of functions that allow range.
            { "STDEV", null },
            { "STDEVA", null },
            { "STDEVP", null },
            { "STDEVPA", null },
            { "SUBTOTAL", Enumerable.Range(1,255).ToList() },
            { "SUM", null },
            { "SUMIF", new List<int> { 0, 2 } },
            { "SUMIFS", new List<int> { 0,1,3,5,7,9} },
            { "SUMPRODUCT", null },
            { "TEXTJOIN", Enumerable.Range(2, 253).ToList() },
            { "VAR", null },
            { "VARA", null },
            { "VARP", null },
            { "VARPA", null },
            { "VLOOKUP", new List<int>{ 1 } },
        };

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
                UnaryOp.SpillRange => throw new NotImplementedException(),
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
                BinaryOp.Intersection => throw new NotImplementedException(),
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
            if (!_functions.TryGetFunc(node.Name, out FormulaFunction function))
            {
                if (!_functions.TryGetFunc(node.Name, out FunctionDefinition legacyFunction))
                    return Error.NameNotRecognized;

                return ExecuteLegacy(context, node, legacyFunction);
            }

            AnyValue?[] args = GetArgs(context, node);
            return function.CallFunction(context, args);
        }

        private AnyValue?[] GetArgs(CalcContext context, FunctionExpression node)
        {
            var args = new AnyValue?[node.Parameters.Count];
            for (var i = 0; i < node.Parameters.Count; ++i)
            {
                var paramNode = node.Parameters[i];
                if (paramNode is not EmptyArgumentNode)
                    args[i] = node.Parameters[i].Accept(context, this);
                else
                    args[i] = null;
            }

            var hasRangeParam = rangeFunctions.TryGetValue(node.Name, out var rangeParamIdx);
            for (var argIdx = 0; argIdx < args.Length; ++argIdx)
            {
                if (!hasRangeParam || (rangeParamIdx is not null && !rangeParamIdx.Contains(argIdx)))
                    args[argIdx] = args[argIdx]?.ImplicitIntersection(context);
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

        private AnyValue ExecuteLegacy(CalcContext context, FunctionExpression node, FunctionDefinition legacyFunction)
        {
            // TODO: Complete rewrite to something sensible
            AnyValue?[] args = GetArgs(context, node);

            // This creates a some of overhead, but all legacy functions will be migrated in near future
            var adaptedArgs = new List<Expression>(args.Length);
            foreach (var arg in args)
            {
                Expression adaptedArg = arg.HasValue ? arg.Value.Match(
                    logical => new Expression(logical),
                    number => new Expression(number),
                    text => new Expression(text),
                    error => new Expression(error),
                    array =>
                    {
                        var castedArray = new double[array.Height, array.Width];
                        for (var row = 0; row < array.Height; ++row)
                            for (var col = 0; col < array.Width; ++col)
                                castedArray[row, col] = array[row, col].Match<double>(
                                    logical => logical ? 1.0 : 0.0,
                                    number => number,
                                    text => throw new NotImplementedException(),
                                    error => throw new NotImplementedException());

                        return new XObjectExpression(castedArray);
                    },
                    range =>
                    {
                        if (range.Areas.Count != 1)
                        {
                            // This sucks. Who ever though it was a good idea to not have reasonable typing system?
                            var references = range.Areas.Select(area =>
                                new CellRangeReference(area.Worksheet.Range(area)));
                            return new XObjectExpression(references);
                        }

                        var area = range.Areas.Single();
                        if (area.Worksheet is not null)
                        {
                            return new XObjectExpression(new CellRangeReference(area.Worksheet.Range(area)));
                        }
                        else
                        {
                            return new XObjectExpression(new CellRangeReference(context.Worksheet.Range(area)));
                        }
                    })
                    : new EmptyValueExpression();
                adaptedArgs.Add(adaptedArg);
            }
            try
            {
                var result = legacyFunction.Function(adaptedArgs);
                return result switch
                {
                    // TODO: Duplicated with logic from CalcEngine
                    bool logic => AnyValue.FromT0(logic),
                    double number => AnyValue.FromT1(number),
                    string text => AnyValue.FromT2(text),
                    int number => AnyValue.FromT1(number), /* date mostly */
                    long number => AnyValue.FromT1(number),
                    DateTime date => AnyValue.FromT1(date.ToOADate()),
                    TimeSpan time => AnyValue.FromT1(time.ToSerialDateTime()),
                    double[,] array => AnyValue.FromT4(new NumberArray(array)),
                    _ => throw new NotImplementedException($"Got a result from some function type {result?.GetType().Name ?? "null"} with value {result}.")
                };
            }
            catch (CalcEngineException)
            {
                //TODO: Map errors to enum
                throw;
            }
        }
    }
}
