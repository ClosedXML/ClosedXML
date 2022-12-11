using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
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
            ce.RegisterFunction("ISLOGICAL", 1, int.MaxValue, IsLogical);
            ce.RegisterFunction("ISNA", 1, int.MaxValue, IsNa);
            ce.RegisterFunction("ISNONTEXT", 1, int.MaxValue, IsNonText);
            ce.RegisterFunction("ISNUMBER", 1, int.MaxValue, IsNumber);
            ce.RegisterFunction("ISODD", 1, 1, Adapt(IsOdd), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("ISREF", 1, int.MaxValue, IsRef);
            ce.RegisterFunction("ISTEXT", 1, int.MaxValue, IsText);
            ce.RegisterFunction("N", 1, N);
            ce.RegisterFunction("NA", 0, NA);
            ce.RegisterFunction("TYPE", 1, Type);
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

        static object IsLogical(List<Expression> p)
        {
            var v = p[0].Evaluate();
            var isLogical = v is bool;

            if (isLogical && p.Count > 1)
            {
                var sublist = p.GetRange(1, p.Count);
                isLogical = (bool)IsLogical(sublist);
            }

            return isLogical;
        }

        static object IsNa(List<Expression> p)
        {
            try
            {
                var v = p[0].Evaluate();
                return v is XLError error && error == XLError.NoValueAvailable;
            }
            catch (NoValueAvailableException)
            {
                return true;
            }
            catch (CalcEngineException)
            {
                return false;
            }
        }

        static object IsNonText(List<Expression> p)
        {
            return !(bool)IsText(p);
        }

        static object IsNumber(List<Expression> p)
        {
            var v = p[0].Evaluate();

            var isNumber = v is double; //Normal number formatting
            if (!isNumber)
            {
                isNumber = v is DateTime; //Handle DateTime Format
            }
            if (!isNumber)
            {
                //Handle Number Styles
                try
                {
                    var stringValue = (string)v;
                    return double.TryParse(stringValue.TrimEnd('%', ' '), NumberStyles.Any, null, out double dv);
                }
                catch (Exception)
                {
                    isNumber = false;
                }
            }

            if (isNumber && p.Count > 1)
            {
                var sublist = p.GetRange(1, p.Count);
                isNumber = (bool)IsNumber(sublist);
            }

            return isNumber;
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

        static object IsRef(List<Expression> p)
        {
            var oe = p[0] as XObjectExpression;
            if (oe == null)
                return false;

            var crr = oe.Value as CellRangeReference;

            return crr != null;
        }

        static object IsText(List<Expression> p)
        {
            //Evaluate Expressions
            var isText = !(bool)string.IsNullOrEmpty(p[0]);
            if (isText)
            {
                isText = !(bool)IsNumber(p);
            }
            if (isText)
            {
                isText = !(bool)IsLogical(p);
            }
            return isText;
        }

        static object N(List<Expression> p)
        {
            return (double)p[0];
        }

        static object NA(List<Expression> p)
        {
            return XLError.NoValueAvailable;
        }

        static object Type(List<Expression> p)
        {
            if ((bool)IsNumber(p))
            {
                return 1;
            }
            if ((bool)IsText(p))
            {
                return 2;
            }
            if ((bool)IsLogical(p))
            {
                return 4;
            }
            if (p[0].Evaluate() is XLError)
            {
                return 16;
            }
            if (p.Count > 1)
            {
                return 64;
            }
            return null;
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
