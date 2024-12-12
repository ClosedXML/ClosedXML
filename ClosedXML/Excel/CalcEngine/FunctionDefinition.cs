using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Function definition class (keeps function name, parameter counts, and delegate).
    /// </summary>
    internal class FunctionDefinition
    {
        private readonly CalcEngineFunction _function;

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

        public int MinParams { get; }

        public int MaxParams { get; }

        public AnyValue CallFunction(CalcContext ctx, Span<AnyValue> args)
        {
            if (ctx.UseImplicitIntersection)
                IntersectArguments(ctx, args);

            return _function(ctx, args);
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

                    var itemResult = _function(ctx, args);

                    // Even if function returns an array, only the top-left value of array is used
                    // as a result for the item, per tests with FILTERXML.
                    result[row, column] = itemResult.TryPickSingleOrMultiValue(out var scalarResult, out var arrayResult, ctx)
                        ? scalarResult
                        : arrayResult[0, 0];
                }
            }

            return new ConstArray(result);
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
