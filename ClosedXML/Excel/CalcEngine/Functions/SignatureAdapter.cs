#nullable disable

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

        public static CalcEngineFunction Adapt(Func<double, double, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToNumber(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                return f(arg0, arg1);
            };
        }

        public static CalcEngineFunction Adapt(Func<double, double, double, bool, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToNumber(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                var arg2Converted = ToNumber(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                    return err2;

                var arg3Converted = CoerceToLogical(args[3], ctx);
                if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                    return err3;

                return f(arg0, arg1, arg2, arg3);
            };
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, double, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0 = args[0];

                var arg1Converted = ToNumber(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                return f(ctx, arg0, arg1);
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
                    return err1;

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

        public static CalcEngineFunction Adapt(Func<CalcContext, double, AnyValue[], AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var argsLoop = args[1..].ToArray();
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

        public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, ScalarValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = args.Length > 1 ? ToScalarValue(args[1], ctx) : ScalarValue.Blank;
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                return f(ctx, arg0, arg1).ToAnyValue();
            };
        }

        public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, AnyValue, ScalarValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToScalarValue(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                var arg2 = args.Length > 2 ? args[2] : AnyValue.Blank;

                return f(ctx, arg0, arg1, arg2).ToAnyValue();
            };
        }

        /// <summary>
        /// Adapt a function that accepts areas as arguments (e.g. SUMPRODUCT). The key benefit is
        /// that all <c>ReferenceArray</c> allocation is done once for a function. The method
        /// shouldn't be used for functions that accept 3D references (e.g. SUMSQ). It is still
        /// necessary to check all errors in the <paramref name="f"/>, adapt method doesn't do that
        /// on its own (potential performance problem). The signature uses an array instead of
        /// IReadOnlyList interface for performance reasons (can't JIT access props through interface).
        /// </summary>
        public static CalcEngineFunction Adapt(Func<CalcContext, Array[], AnyValue> f)
        {
            return (ctx, args) =>
            {
                var areas = new Array[args.Length];
                for (var i = 0; i < args.Length; ++i)
                {
                    areas[i] = args[i].TryPickSingleOrMultiValue(out var scalar, out var array, ctx)
                        ? new ScalarArray(scalar, 1, 1)
                        : array;
                }

                return f(ctx, areas);
            };
        }

        public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, AnyValue, double, bool, AnyValue> f, bool defaultValue0)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToScalarValue(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1 = args[1];

                var arg2Converted = ToNumber(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                    return err2;

                var arg3Converted = args.Length >= 4 ? CoerceToLogical(args[3], ctx) : defaultValue0;
                if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                    return err3;

                return f(ctx, arg0, arg1, arg2, arg3);
            };
        }

        public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, double, double, AnyValue> f, double defaultValue0, double defaultValue1)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToNumber(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                var arg2Converted = ToNumber(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                    return err2;

                var arg3Optional = defaultValue0;
                if (args.Length >= 4)
                {
                    var arg3Converted = ToNumber(args[3], ctx);
                    if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                        return err3;

                    arg3Optional = arg3;
                }

                var arg4Optional = defaultValue1;
                if (args.Length >= 5)
                {
                    var arg4Converted = ToNumber(args[4], ctx);
                    if (!arg4Converted.TryPickT0(out var arg4, out var err4))
                        return err4;

                    arg4Optional = arg4;
                }

                return f(arg0, arg1, arg2, arg3Optional, arg4Optional);
            };
        }

        public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, double, double, double, AnyValue> f, double defaultValue0, double defaultValue1)
        {
            return (ctx, args) =>
            {
                var arg0Converted = ToNumber(args[0], ctx);
                if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                    return err0;

                var arg1Converted = ToNumber(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                    return err1;

                var arg2Converted = ToNumber(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                    return err2;

                var arg3Converted = ToNumber(args[3], ctx);
                if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                    return err3;

                var arg4Optional = defaultValue0;
                if (args.Length >= 5)
                {
                    var arg4Converted = ToNumber(args[4], ctx);
                    if (!arg4Converted.TryPickT0(out var arg4, out var err4))
                        return err4;

                    arg4Optional = arg4;
                }

                var arg5Optional = defaultValue1;
                if (args.Length >= 6)
                {
                    var arg5Converted = ToNumber(args[5], ctx);
                    if (!arg5Converted.TryPickT0(out var arg5, out var err5))
                        return err5;

                    arg5Optional = arg5;
                }

                return f(arg0, arg1, arg2, arg3, arg4Optional, arg5Optional);
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
