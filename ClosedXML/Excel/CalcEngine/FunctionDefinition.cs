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
        /// <summary>
        /// Only the <see cref="_function"/> or <see cref="_legacyFunction"/> is set.
        /// </summary>
        private readonly CalcEngineFunction _function;

        /// <summary>
        /// Only the <see cref="_function"/> or <see cref="_legacyFunction"/> is set.
        /// </summary>
        private readonly LegacyCalcEngineFunction _legacyFunction;

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

            if (_legacyFunction is not null)
            {
                // This creates a some of overhead, but all legacy functions will be migrated in near future
                var adaptedArgs = new List<Expression>(args.Length);
                foreach (var arg in args)
                    adaptedArgs.Add(ConvertAnyValueToLegacyExpression(ctx, arg));

                var result = _legacyFunction(adaptedArgs);
                return ConvertLegacyFormulaValueToAnyValue(result);
            }

            return _function(ctx, args);
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

        public static AnyValue ConvertLegacyFormulaValueToAnyValue(object result)
        {
            return result switch
            {
                bool logic => AnyValue.From(logic),
                double number => AnyValue.From(number),
                string text => AnyValue.From(text),
                int number => AnyValue.From(number), /* int represents a date in most cases (legacy functions, e.g. SECOND), number are double */
                long number => AnyValue.From(number),
                DateTime date => AnyValue.From(date.ToOADate()),
                TimeSpan time => AnyValue.From(time.ToSerialDateTime()),
                XLError errorType => AnyValue.From(errorType),
                double[,] array => AnyValue.From(new NumberArray(array)),
                _ => throw new NotImplementedException($"Got a result from some function type {result?.GetType().Name ?? "null"} with value {result}.")
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
                        for (var col = 0; col < array.Width; ++col)
                            convertedArray[row, col] = array[row, col].Match(
                                () => 0.0,
                                logical => logical ? 1.0 : 0.0,
                                number => number,
                                text => throw new NotImplementedException(),
                                error => throw new NotImplementedException());

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
    }
}
