using System;
using System.Collections.Generic;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class Information
    {
        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("ERROR.TYPE", 1, 1, Adapt(ErrorType), FunctionFlags.Scalar);
            ce.RegisterFunction("ISBLANK", 1, 1, Adapt(IsBlank), FunctionFlags.Scalar);
            ce.RegisterFunction("ISERR", 1, 1, Adapt(IsErr), FunctionFlags.Scalar);
            ce.RegisterFunction("ISERROR", 1, 1, Adapt(IsError), FunctionFlags.Scalar);
            ce.RegisterFunction("ISEVEN", 1, 1, Adapt(IsEven), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("ISLOGICAL", 1, 1, Adapt(IsLogical), FunctionFlags.Scalar);
            ce.RegisterFunction("ISNA", 1, 1, Adapt(IsNa), FunctionFlags.Scalar);
            ce.RegisterFunction("ISNONTEXT", 1, 1, Adapt(IsNonText), FunctionFlags.Scalar);
            ce.RegisterFunction("ISNUMBER", 1, 1, Adapt(IsNumber), FunctionFlags.Scalar);
            ce.RegisterFunction("ISODD", 1, 1, Adapt(IsOdd), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("ISREF", 1, 1, Adapt(IsRef), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("ISTEXT", 1, 1, Adapt(IsText), FunctionFlags.Scalar);
            ce.RegisterFunction("N", 1, N);
            ce.RegisterFunction("NA", 0, NA);
            ce.RegisterFunction("TYPE", 1, 1, Adapt(Type), FunctionFlags.Range, AllowRange.All);
        }

        private static AnyValue ErrorType(CalcContext ctx, ScalarValue value)
        {
            if (!value.TryPickError(out var error))
                return XLError.NoValueAvailable;

            return error switch
            {
                XLError.NullValue => 1,
                XLError.DivisionByZero => 2,
                XLError.IncompatibleValue => 3,
                XLError.CellReference => 4,
                XLError.NameNotRecognized => 5,
                XLError.NumberInvalid => 6,
                XLError.NoValueAvailable => 7,
                _ => throw new NotSupportedException($"Error {error} not supported.")
            };
        }

        private static AnyValue IsBlank(CalcContext ctx, ScalarValue value)
        {
            return value.IsBlank;
        }

        private static AnyValue IsErr(CalcContext ctx, ScalarValue value)
        {
            return value.TryPickError(out var error) && error != XLError.NoValueAvailable;
        }

        private static AnyValue IsError(CalcContext ctx, ScalarValue value)
        {
            return value.TryPickError(out _);
        }

        private static AnyValue IsEven(CalcContext ctx, AnyValue value)
        {
            return GetParity(ctx, value, static (scalar, ctx) =>
            {
                if (scalar.IsLogical)
                    return XLError.IncompatibleValue;

                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                    return error;

                return Math.Truncate(number) % 2 == 0;
            });
        }

        private static AnyValue IsLogical(CalcContext ctx, ScalarValue value)
        {
            return value.IsLogical;
        }

        private static AnyValue IsNa(CalcContext ctx, ScalarValue value)
        {
            return value.TryPickError(out var error) && error == XLError.NoValueAvailable;
        }

        private static AnyValue IsNonText(CalcContext ctx, ScalarValue value)
        {
            return !value.IsText;
        }

        private static AnyValue IsNumber(CalcContext ctx, ScalarValue value)
        {
            return value.IsNumber;
        }

        private static AnyValue IsOdd(CalcContext ctx, AnyValue value)
        {
            return GetParity(ctx, value, static (scalar, ctx) =>
            {
                if (scalar.IsLogical)
                    return XLError.IncompatibleValue;

                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                    return error;

                return Math.Truncate(number) % 2 != 0;
            });
        }

        private static AnyValue IsRef(CalcContext ctx, AnyValue value)
        {
            return value.IsReference;
        }

        private static AnyValue IsText(CalcContext ctx, ScalarValue value)
        {
            return value.IsText;
        }
        
        static object N(List<Expression> p)
        {
            return (double)p[0];
        }

        static object NA(List<Expression> p)
        {
            return XLError.NoValueAvailable;
        }

        private static AnyValue Type(CalcContext ctx, AnyValue value)
        {
            if (!value.TryPickScalar(out var scalar, out var collection))
            {
                var isArray = collection.TryPickT0(out _, out var reference);
                if (isArray)
                    return 64;
                if (reference.Areas.Count > 1)
                    return 16;
                if (!reference.TryGetSingleCellValue(out scalar, ctx))
                    return 64;
            }

            if (scalar.IsBlank || scalar.IsNumber)
                return 1;
            if (scalar.IsText)
                return 2;
            if (scalar.IsLogical)
                return 4;
            if (scalar.IsError)
                return 16;

            // There is a "composite type", but no idea what exactly it is. Shouldn't happen.
            throw new InvalidOperationException("Unknown type.");
        }

        private static AnyValue GetParity(CalcContext ctx, AnyValue value, Func<ScalarValue, CalcContext, ScalarValue> f)
        {
            // IsOdd/IsEven has very strange semantic that is different for pretty much every other function
            // Array behaves differently for multi-cell references, in-place blank vs cell blank give different value...
            if (value.TryPickScalar(out var scalar, out var coll))
            {
                if (scalar.IsBlank)
                    return XLError.NoValueAvailable;

                return f(scalar, ctx).ToAnyValue();
            }

            if (coll.TryPickT0(out var array, out var reference))
                return array.Apply(x => f(x, ctx));

            if (!reference.TryGetSingleCellValue(out var cellValue, ctx))
                return XLError.IncompatibleValue;

            return f(cellValue, ctx).ToAnyValue();
        }
    }
}
