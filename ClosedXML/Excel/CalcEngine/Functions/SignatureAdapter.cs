using System;
using System.Collections.Generic;

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
                if (!ctx.Converter.ToText(args[0] ?? AnyValue.From((double)0)).TryPickT0(out var arg0, out var error))
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
                    if (!(args[i] ?? 0).TryPickReference(out var reference))
                        return Error.CellValue;

                    references.Add(reference);
                }

                return f(ctx, number, references);
            };
        }
    }
}
