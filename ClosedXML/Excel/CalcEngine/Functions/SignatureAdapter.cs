using System;
using System.Collections.Generic;

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

        public static CalcEngineFunction Adapt(Func<AnyValue> f)
        {
            return (_, _) => f();
        }

        public static CalcEngineFunction AdaptCoerced(Func<Boolean, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = CoerceToLogical(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                return f(arg0);
            };
        }

        public static CalcEngineFunction Adapt(Func<double, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                return f(arg0);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue?, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToText(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1 = default(ScalarValue?);
                if (args.Length > 1)
                {
                    var arg1Converted = ToScalarValue(args[1], ctx);
                    if (!arg1Converted.TryPickT0(out var arg1Value, out var err1))
                        return err1;

                    arg1 = arg1Value;
                }


                return f(ctx, arg0, arg1);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, AnyValue> f)
        {
            return (ctx, args) => f(ctx, args[0]);
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                return f(ctx, arg0);
            };
        }

        public static CalcEngineFunction Adapt(Func<ScalarValue, ScalarValue, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToScalarValue(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err0;

                return f(arg0, arg1);
            };
        }

        public static CalcEngineFunction AdaptLastOptional(Func<ScalarValue, AnyValue, AnyValue, AnyValue> f, AnyValue lastDefault)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1 = args[1];
                var arg2 = args.Length > 2 ? args[2] : lastDefault;
                return f(arg0, arg1, arg2);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, double, List<Reference>, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var argsLoop = new List<Reference>();
                for (var i = 1; i < args.Length; ++i)
                {
                    if (!args[i].TryPickReference(out var reference, out var error))
                        return error;

                    argsLoop.Add(reference);
                }

                return f(ctx, arg0, argsLoop);
            };
        }

        public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, string, string, OneOf<double, Blank>, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToText(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToText(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                OneOf<double, Blank> arg2Optional = Blank.Value;
                if (args.Length > 2)
                {
                    var arg2Converted = ToNumber(args[2], ctx);
                    if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                        return err2;

                    arg2Optional = arg2;
                }

                return f(ctx, arg0, arg1, arg2Optional);
            };
        }

        public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, AnyValue, ScalarValue, ScalarValue, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1 = args[1];

                var arg2Converted = ToScalarValue(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                    return err2;

                var arg3Converted = args.Length >= 4 ? ToScalarValue(args[3], ctx) : ScalarValue.Blank;
                if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                    return err3;

                return f(ctx, arg0, arg1, arg2, arg3);
            };
        }

        #endregion

        #region Value converters
        // Each method is named ToSomething and it converts an argument into a desired type (e.g. for ToSomething it should be type Something).
        // Return value is always OneOf<Something, Error>, if there is an error, return it as an error.

        private static OneOf<Boolean, XLError> CoerceToLogical(in AnyValue value, CalcContext ctx)
        {
            if (!ToScalarValue(in value, ctx).TryPickT0(out var scalar, out var scalarError))
                return scalarError;

            if (!scalar.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var coercionError))
                return coercionError;

            return logical;
        }

        private static OneOf<double, XLError> ToNumber(in AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToNumber(ctx.Culture);

            if (collection.TryPickT0(out _, out var reference))
                throw new NotImplementedException("Array formulas not implemented.");

            if (reference.TryGetSingleCellValue(out var scalarValue, ctx))
                return scalarValue.ToNumber(ctx.Culture);

            throw new NotImplementedException("Array formulas not implemented.");
        }

        private static OneOf<string, XLError> ToText(in AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToText(ctx.Culture);

            if (collection.TryPickT0(out _, out var reference))
                throw new NotImplementedException("Array formulas not implemented.");

            if (reference.TryGetSingleCellValue(out var scalarValue, ctx))
                return scalarValue.ToText(ctx.Culture);

            throw new NotImplementedException("Array formulas not implemented.");
        }

        private static OneOf<ScalarValue, XLError> ToScalarValue(in AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar;

            if (collection.TryPickT0(out var array, out var reference))
                return array[0, 0];

            if (reference.TryGetSingleCellValue(out var referenceScalar, ctx))
                return referenceScalar;

            return OneOf<ScalarValue, XLError>.FromT1(XLError.IncompatibleValue);
        }

        #endregion
    }
}
