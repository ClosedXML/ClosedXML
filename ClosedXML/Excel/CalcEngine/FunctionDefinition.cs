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
        public FunctionDefinition(string name, int minParams, int maxParams, CalcEngineFunction function, FunctionFlags flags, AllowRange allowRanges, IReadOnlyCollection<int> markedParams)
        {
            if (allowRanges == AllowRange.None && markedParams.Any())
                throw new ArgumentException(nameof(markedParams));

            Name = name;
            MinParams = minParams;
            MaxParams = maxParams;
            AllowRanges = allowRanges;
            MarkedParams = markedParams;
            Function = function;
            Flags = flags;
        }

        public FunctionDefinition(string name, int minParams, int maxParams, LegacyCalcEngineFunction function, AllowRange allowRanges, IReadOnlyCollection<int> markedParams)
        {
            if (allowRanges == AllowRange.None && markedParams.Any())
                throw new ArgumentException(nameof(markedParams));

            Name = name;
            MinParams = minParams;
            MaxParams = maxParams;
            AllowRanges = allowRanges;
            MarkedParams = markedParams;
            LegacyFunction = function;
        }

        public string Name { get; }

        public int MinParams { get; }

        public int MaxParams { get; }

        public FunctionFlags Flags { get; }

        public AllowRange AllowRanges { get; }

        public IReadOnlyCollection<int> MarkedParams { get; }

        /// <summary>Only the <see cref="Function"/> or <see cref="LegacyFunction"/> is set.</summary>
        public CalcEngineFunction Function { get; }

        /// <summary>Only the <see cref="Function"/> or <see cref="LegacyFunction"/> is set.</summary>
        public LegacyCalcEngineFunction LegacyFunction { get; }

        public AnyValue CallFunction(CalcContext ctx, params AnyValue?[] args)
        {
            if (LegacyFunction is not null)
            {
                // This creates a some of overhead, but all legacy functions will be migrated in near future
                var adaptedArgs = new List<Expression>(args.Length);
                foreach (var arg in args)
                    adaptedArgs.Add(ConvertAnyValueToLegacyExpression(ctx, arg));

                var result = LegacyFunction(adaptedArgs);
                return ConvertLegacyFormulaValueToAnyValue(result);
            }

            return Function(ctx, args);
        }

        public static AnyValue ConvertLegacyFormulaValueToAnyValue(object result)
        {
            return result switch
            {
                bool logic => AnyValue.FromT0(logic),
                double number => AnyValue.FromT1(number),
                string text => AnyValue.FromT2(text),
                int number => AnyValue.FromT1(number), /* int represents a date in most cases (legacy functions, e.g. SECOND), number are double */
                long number => AnyValue.FromT1(number),
                DateTime date => AnyValue.FromT1(date.ToOADate()),
                TimeSpan time => AnyValue.FromT1(time.ToSerialDateTime()),
                Error errorType => AnyValue.FromT3(errorType),
                double[,] array => AnyValue.FromT4(new NumberArray(array)),
                _ => throw new NotImplementedException($"Got a result from some function type {result?.GetType().Name ?? "null"} with value {result}.")
            };
        }

        private static Expression ConvertAnyValueToLegacyExpression(CalcContext context, AnyValue? arg)
        {
            if (!arg.HasValue)
                return new EmptyValueExpression();

            return arg.Value.Match(
                logical => new Expression(logical),
                number => new Expression(number),
                text => new Expression(text),
                error => new Expression(error),
                array =>
                {
                    var castedArray = new double[array.Height, array.Width];
                    for (var row = 0; row < array.Height; ++row)
                        for (var col = 0; col < array.Width; ++col)
                            castedArray[row, col] = array[row, col].Match(
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
                            new CellRangeReference((area.Worksheet ?? context.Worksheet).Range(area)));
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
                });
        }
    }
}
