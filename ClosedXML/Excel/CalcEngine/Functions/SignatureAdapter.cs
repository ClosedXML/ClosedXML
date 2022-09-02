using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Presentation;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    /// <summary>
    /// A collection of adapter functions from a more a generic formula function to more specific ones.
    /// </summary>
    internal static class SignatureAdapter
    {
        #region Signature adapters
        // Each method converts a more specific signature of a function into a generic formula function type.
        // We have many functions with same signature and the adapters should be reusable. Convert parameters
        // through value converters below. We can hopefully generate them at a later date, so try to keep them similar.

        public static CalcEngineFunction Adapt(Func<double, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Input = args[0] ?? 0;
                var arg0Converted = ToNumber(arg0Input, ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                return f(arg0);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue?, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Input = args[0] ?? 0;
                var arg0Converted = ToText(arg0Input, ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1 = default(ScalarValue?);
                if (args.Length > 1)
                {
                    var arg1Input = args[1] ?? 0;
                    var arg1Converted = ToScalarValue(arg1Input, ctx);
                    if (!arg1Converted.TryPickT0(out var arg1Value, out var err1))
                        return err1;

                    arg1 = arg1Value;
                }


                return f(ctx, arg0, arg1);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, double, List<Reference>, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Input = args[0] ?? 0;
                var arg0Converted = ToNumber(arg0Input, ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var argsLoop = new List<Reference>();
                for (var i = 1; i < args.Length; ++i)
                {
                    var argLoopInput = args[i] ?? 0;
                    if (!argLoopInput.TryPickReference(out var reference))
                        return Error.CellValue;

                    argsLoop.Add(reference);
                }

                return f(ctx, arg0, argsLoop);
            };
        }

        #endregion

        #region Value converters
        // Each method is named ToSomething and it converts an argument into a desired type (e.g. for ToSomething it should be type Something).
        // Return value is always OneOf<Something, Error>, if there is an error, return it as an error.
        
        private static OneOf<double, Error> ToNumber(in AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToNumber(ctx.Culture);

            return collection.Match(
                _ => throw new NotImplementedException("Reference to number conversion not yet implemented."),
                reference =>
                {
                    if (reference.TryGetSingleCellValue(out var scalarValue, ctx))
                        return scalarValue.ToNumber(ctx.Culture);

                    throw new NotImplementedException("Not sure what to do with it.");
                });
        }

        private static OneOf<string, Error> ToText(in AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToText(ctx.Culture);

            if (collection.TryPickT0(out var array, out var _))
                return array[0, 0].ToText(ctx.Culture);

            throw new NotImplementedException("Conversion from reference to text is not implemented yet.");
        }

        private static OneOf<ScalarValue, Error> ToScalarValue(in AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar;

            if (collection.TryPickT0(out var array, out var reference))
                return array[0, 0];

            if (reference.TryGetSingleCellValue(out var referenceScalar, ctx))
                return referenceScalar;

            return OneOf<ScalarValue, Error>.FromT1(Error.CellValue);
        }

        #endregion
    }
}
