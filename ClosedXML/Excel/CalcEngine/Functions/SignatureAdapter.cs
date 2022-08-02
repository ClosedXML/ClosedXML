using System;
using System.Collections.Generic;
using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error>;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    /// <summary>
    /// A collection of adapter functions from a more a generic formula function to more specific ones.
    /// </summary>
    internal static class SignatureAdapter
    {
        public static CalcEngineFunction Adapt(Func<CalcContext, string, AnyValue?, AnyValue> f)
        {
            return (ctx, args) =>
            {
                if (!ctx.Converter.ToText(args[0] ?? AnyValue.FromT1(0)).TryPickT0(out var arg0, out var error))
                    return error;

                return f(ctx, arg0, args.Length > 1 ? args[1] : null);
            };
        }

        public static CalcEngineFunction Adapt(Func<double, AnyValue> f)
        {
            return (ctx, args) => ctx.Converter.ToNumber(args[0]).Match(
                    number => f(number),
                    error => error);
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, double, List<Reference>, AnyValue> f)
        {
            return (ctx, args) =>
            {
                if (!ctx.Converter.ToNumber(args[0] ?? 0).TryPickT0(out var number, out var error))
                    return error;

                var references = new List<Reference>();
                for (var i = 1; i < args.Length; ++i)
                {
                    if (!(args[i] ?? 0).TryPickT5(out var reference, out var rest))
                        return Error.CellValue;

                    references.Add(reference);
                }

                return f(ctx, number, references);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue?, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0 = args[0];
                var convertedArg0 = arg0.HasValue
                    ? ConvertToScalar(ctx, arg0.Value)
                    : null;

                return f(ctx, convertedArg0);
            };
        }

        private static ScalarValue ConvertToScalar(CalcContext ctx, AnyValue val)
        {
            if (val.TryPickScalar(out var scalar, out var collection))
                return scalar;

            return collection.Match(
                array => array[0, 0],
                reference =>
                {
                    if (!reference.TryGetSingleCellValue(out var scalar, ctx))
                        return scalar;

                        // This should never happen:
                        // * Param is a scalar and arg is a multi-cell range - implicit intersection turns it to single-cell reference/error
                        // * Param is a range and arg is a multi-cell range and the Adapt calls this function that converts to range to scalar - doesn't make sense.
                    throw new InvalidOperationException("Trying to convert multi-cell reference to a scalar has unknown semantic.");
                });
        }
    }
}
