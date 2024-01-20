using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Exceptions;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Function definition class (keeps function name, parameter counts, and delegate).
    /// </summary>
    internal class FunctionDefinition
    {
        /// <summary>
        /// Only the <see cref="_function"/> or <see cref="_legacyFunction"/> is set.
        /// </summary>
        private readonly CalcEngineFunction? _function;

        /// <summary>
        /// Only the <see cref="_function"/> or <see cref="_legacyFunction"/> is set.
        /// </summary>
        private readonly LegacyCalcEngineFunction? _legacyFunction;

        private readonly FunctionFlags _flags;

        private readonly AllowRange _allowRanges;

        /// <summary>
        /// Which parameters of the function are marked. The values are indexes of the function parameters, starting from 0.
        /// Used to determine which arguments allow ranges and which don't.
        /// </summary>
        private readonly IReadOnlyCollection<int> _markedParams;

        public FunctionDefinition(int minParams, int maxParams, CalcEngineFunction function, FunctionFlags flags, AllowRange allowRanges, IReadOnlyCollection<int> markedParams)
        {
            if (allowRanges == AllowRange.None && markedParams.Any())
                throw new ArgumentException(nameof(markedParams));

            MinParams = minParams;
            MaxParams = maxParams;
            _allowRanges = allowRanges;
            _markedParams = markedParams;
            _function = function;
            _flags = flags;
        }

        public FunctionDefinition(int minParams, int maxParams, LegacyCalcEngineFunction function, AllowRange allowRanges, IReadOnlyCollection<int> markedParams)
        {
            if (allowRanges == AllowRange.None && markedParams.Any())
                throw new ArgumentException(nameof(markedParams));

            MinParams = minParams;
            MaxParams = maxParams;
            _allowRanges = allowRanges;
            _markedParams = markedParams;
            _legacyFunction = function;
        }

        public int MinParams { get; }

        public int MaxParams { get; }

        public AnyValue CallFunction(CalcContext ctx, Span<AnyValue> args)
        {
            if (ctx.UseImplicitIntersection)
                IntersectArguments(ctx, args);

            return EvaluateFunction(ctx, args);
        }

        /// <summary>
        /// Evaluate the function with array formula semantic.
        /// </summary>
        public AnyValue CallAsArray(CalcContext ctx, Span<AnyValue> args)
        {
            if (_flags.HasFlag(FunctionFlags.ReturnsArray) && _allowRanges == AllowRange.All)
            {
                return _function!(ctx, args);
            }

            // Step 1: For scalar parameters of function, determine maximum size of scalar
            // parameters from argument arrays
            var (totalRows, totalColumns) = GetScalarArgsMaxSize(args);

            // Step 2: Normalize arguments. Single params are converted to array of same size, multi params are converted from scalars
            for (var i = 0; i < args.Length; ++i)
            {
                ref var arg = ref args[i];
                var argIsSingle = arg.TryPickSingleOrMultiValue(out var single, out var multi, ctx);
                if (IsParameterSingleValue(i))
                {
                    arg = argIsSingle
                        ? new ScalarArray(single, totalColumns, totalRows)
                        : multi.Broadcast(totalRows, totalColumns);
                }
                else
                {
                    // 18.17.2.4 When a function expects a multi-valued argument but a single-valued
                    // expression is passed, that single-valued argument is treated as a 1x1 array.
                    // If there is an error as a single value, e.g. reference to a single cell, the SUMIF behaves
                    // as it was converted to 1x1 array and doesn't return error, just because it found an error.
                    // Ergo: for ranges, we don't immediately return error, just because range parameter contains an error
                    arg = argIsSingle
                        ? new ScalarArray(single, 1, 1)
                        : multi;
                }
            }

            // Step 3: For each item in total array, calculate function
            var result = new ScalarValue[totalRows, totalColumns];
            for (var row = 0; row < totalRows; ++row)
            {
                for (var column = 0; column < totalColumns; ++column)
                {
                    var itemArg = new AnyValue[args.Length];
                    for (var i = 0; i < itemArg.Length; ++i)
                    {
                        ref var arg = ref args[i];
                        itemArg[i] = IsParameterSingleValue(i)
                            ? arg.GetArray()[row, column].ToAnyValue()
                            : arg;
                    }

                    var itemResult = EvaluateFunction(ctx, itemArg);

                    // Even if function returns an array, only the top-left value of array is used
                    // as a result for the item, per tests with FILTERXML.
                    result[row, column] = itemResult.TryPickSingleOrMultiValue(out var scalarResult, out var arrayResult, ctx)
                        ? scalarResult
                        : arrayResult[0, 0];
                }
            }

            return new ConstArray(result);
        }

        private AnyValue EvaluateFunction(CalcContext ctx, Span<AnyValue> args)
        {
            if (_legacyFunction is not null)
            {
                // This creates a some of overhead, but all legacy functions will be migrated in near future
                var adaptedArgs = new List<Expression>(args.Length);
                foreach (var arg in args)
                    adaptedArgs.Add(ConvertAnyValueToLegacyExpression(ctx, arg));

                try
                {
                    var result = _legacyFunction(adaptedArgs);
                    return ConvertLegacyFormulaValueToAnyValue(result);
                }
                catch (FormulaErrorException ex)
                {
                    return ex.Error;
                }
            }

            return _function!(ctx, args);
        }

        private void IntersectArguments(CalcContext ctx, Span<AnyValue> args)
        {
            for (var i = 0; i < args.Length; ++i)
            {
                var intersectArgument = _allowRanges switch
                {
                    AllowRange.None => true,
                    AllowRange.Except => _markedParams.Contains(i),
                    AllowRange.Only => !_markedParams.Contains(i),
                    AllowRange.All => false,
                    _ => throw new InvalidOperationException($"Unexpected value {_allowRanges}")
                };
                if (intersectArgument)
                    args[i] = args[i].ImplicitIntersection(ctx);
            }
        }

        private static AnyValue ConvertLegacyFormulaValueToAnyValue(object? result)
        {
            return result switch
            {
                null => AnyValue.Blank,
                bool logic => AnyValue.From(logic),
                double number => AnyValue.From(number),
                string text => AnyValue.From(text),
                int number => AnyValue.From(number), /* int represents a date in most cases (legacy functions, e.g. SECOND), number are double */
                long number => AnyValue.From(number),
                DateTime date => AnyValue.From(date.ToOADate()),
                TimeSpan time => AnyValue.From(time.ToSerialDateTime()),
                XLError errorType => AnyValue.From(errorType),
                double[,] array => AnyValue.From(new NumberArray(array)),
                XLCellValue cellValue => ((ScalarValue)cellValue).ToAnyValue(), // Some functions return directly value of a cell.
                _ => throw new NotImplementedException($"Got a result from some function type {result.GetType().Name} with value {result}.")
            };
        }

        private static Expression ConvertAnyValueToLegacyExpression(CalcContext context, AnyValue arg)
        {
            return arg.Match(
                () => new EmptyValueExpression(),
                logical => new Expression(logical),
                number => new Expression(number),
                text => new Expression(text),
                error => new Expression(error),
                array =>
                {
                    var convertedArray = new double[array.Height, array.Width];
                    for (var row = 0; row < array.Height; ++row)
                    {
                        for (var col = 0; col < array.Width; ++col)
                        {
                            var value = array[row, col];

                            // Generally speaking, once a value in a parameter is an error,
                            // function returns that error. There are few outliers, but very rare.
                            // Since legacy function never supported errors, use the expression
                            // that throws on error which is caught by legacy function evaluator.
                            // That will simulate the correct behavior in most cases:
                            // * OK scalar legacy function that is applied one by one
                            // * OK reducing legacy function that maps multiple values to one value.
                            // * NOK function returning array (e.g MATMULT)
                            if (value.IsError)
                                return new Expression(value.GetError());

                            convertedArray[row, col] = value.Match(
                                () => 0.0,
                                logical => logical ? 1.0 : 0.0,
                                number => number,
                                text => throw new NotImplementedException(),
                                _ => throw new UnreachableException());
                        }
                    }

                    return new XObjectExpression(convertedArray);
                },
                range =>
                {
                    if (range.Areas.Count != 1)
                    {
                        var references = range.Areas.Select(area =>
                            new CellRangeReference((area.Worksheet ?? context.Worksheet).Range(area))).ToList();
                        return new XObjectExpression(references);
                    }

                    var area = range.Areas.Single();
                    var ws = area.Worksheet ?? context.Worksheet;
                    return new XObjectExpression(new CellRangeReference(ws.Range(area)));
                });
        }

        private (int Rows, int Columns) GetScalarArgsMaxSize(Span<AnyValue> args)
        {
            var maxRows = 1;
            var maxColumns = 1;
            for (var i = 0; i < args.Length; ++i)
            {
                ref var arg = ref args[i];
                if (IsParameterSingleValue(i))
                {
                    var (argRows, argColumns) = arg.GetArraySize();
                    maxRows = Math.Max(maxRows, argRows);
                    maxColumns = Math.Max(maxColumns, argColumns);
                }
            }

            return (maxRows, maxColumns);
        }

        private bool IsParameterSingleValue(int paramIndex)
        {
            var paramAllowsMultiValues = _allowRanges switch
            {
                AllowRange.None => false,
                AllowRange.Except => !_markedParams.Contains(paramIndex),
                AllowRange.Only => _markedParams.Contains(paramIndex),
                AllowRange.All => true,
                _ => throw new NotSupportedException($"Unexpected value {_allowRanges}")
            };
            return !paramAllowsMultiValues;
        }
    }
}
