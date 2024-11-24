// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class MathTrig
    {
        /// <summary>
        /// Maximum integer number that can be precisely represented in a double.
        /// Calculated as <c>Math.Pow(2, 53) - 1</c>, but use literal to make it
        /// constant (=usable in pattern matching).
        /// </summary>
        private const double MaxDoubleInt = 9007199254740991;

        private static readonly Random _rnd = new Random();

        /// <summary>
        /// Key: roman form. Value: A collection of subtract symbols and subtract value.
        /// Collection is sorted by subtract value in descending order.
        /// </summary>
        private static readonly Lazy<IReadOnlyDictionary<int, IReadOnlyList<(string Symbol, int Value)>>> RomanForms = new(BuildRomanForms);

        private static readonly IReadOnlyDictionary<char, int> RomanSymbolValues = new Dictionary<char, int>
        {
            {'I', 1},
            {'V', 5},
            {'X', 10},
            {'L', 50},
            {'C', 100},
            {'D', 500},
            {'M', 1000}
        };

        #region Register

        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("ABS", 1, 1, Adapt(Abs), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOS", 1, 1, Adapt(Acos), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOSH", 1, 1, Adapt(Acosh), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOT", 1, 1, Adapt(Acot), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("ACOTH", 1, 1, Adapt(Acoth), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("ARABIC", 1, 1, Adapt(Arabic), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("ASIN", 1, 1, Adapt(Asin), FunctionFlags.Scalar);
            ce.RegisterFunction("ASINH", 1, 1, Adapt(Asinh), FunctionFlags.Scalar);
            ce.RegisterFunction("ATAN", 1, 1, Adapt(Atan), FunctionFlags.Scalar);
            ce.RegisterFunction("ATAN2", 2, 2, Adapt(Atan2), FunctionFlags.Scalar);
            ce.RegisterFunction("ATANH", 1, 1, Adapt(Atanh), FunctionFlags.Scalar);
            ce.RegisterFunction("BASE", 2, 3, AdaptLastOptional(Base, 1), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("CEILING", 2, 2, Adapt(Ceiling), FunctionFlags.Scalar);
            ce.RegisterFunction("CEILING.MATH", 1, 3, AdaptLastTwoOptional(CeilingMath, 1, 0), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("COMBIN", 2, 2, Adapt(Combin), FunctionFlags.Scalar);
            ce.RegisterFunction("COMBINA", 2, 2, Adapt(CombinA), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("COS", 1, 1, Adapt(Cos), FunctionFlags.Scalar);
            ce.RegisterFunction("COSH", 1, 1, Adapt(Cosh), FunctionFlags.Scalar);
            ce.RegisterFunction("COT", 1, 1, Adapt(Cot), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("COTH", 1, 1, Adapt(Coth), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("CSC", 1, 1, Adapt(Csc), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("CSCH", 1, 1, Adapt(Csch), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("DECIMAL", 2, 2, Adapt(Decimal), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("DEGREES", 1, 1, Adapt(Degrees), FunctionFlags.Scalar);
            ce.RegisterFunction("EVEN", 1, 1, Adapt(Even), FunctionFlags.Scalar);
            ce.RegisterFunction("EXP", 1, 1, Adapt(Exp), FunctionFlags.Scalar);
            ce.RegisterFunction("FACT", 1, 1, Adapt(Fact), FunctionFlags.Scalar);
            ce.RegisterFunction("FACTDOUBLE", 1, 1, Adapt(FactDouble), FunctionFlags.Scalar);
            ce.RegisterFunction("FLOOR", 2, 2, Adapt(Floor), FunctionFlags.Scalar);
            ce.RegisterFunction("FLOOR.MATH", 1, 3, AdaptLastTwoOptional(FloorMath, 1, 0), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("GCD", 1, 255, Adapt(Gcd), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("INT", 1, 1, Adapt(Int), FunctionFlags.Scalar);
            ce.RegisterFunction("LCM", 1, 255, Adapt(Lcm), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("LN", 1, 1, Adapt(Ln), FunctionFlags.Scalar);
            ce.RegisterFunction("LOG", 1, 2, AdaptLastOptional(Log, 10), FunctionFlags.Scalar);
            ce.RegisterFunction("LOG10", 1, 1, Adapt(Log10), FunctionFlags.Scalar);
            ce.RegisterFunction("MDETERM", 1, 1, Adapt(MDeterm), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("MINVERSE", 1, 1, Adapt(MInverse), FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All);
            ce.RegisterFunction("MMULT", 2, MMult, AllowRange.All);
            ce.RegisterFunction("MOD", 2, 2, Adapt(Mod), FunctionFlags.Scalar);
            ce.RegisterFunction("MROUND", 2, 2, Adapt(MRound), FunctionFlags.Scalar);
            ce.RegisterFunction("MULTINOMIAL", 1, 255, AdaptMultinomial(Multinomial), FunctionFlags.Scalar, AllowRange.All);
            ce.RegisterFunction("ODD", 1, 1, Adapt(Odd), FunctionFlags.Scalar);
            ce.RegisterFunction("PI", 0, 0, Adapt(Pi), FunctionFlags.Scalar);
            ce.RegisterFunction("POWER", 2, 2, Adapt(Power), FunctionFlags.Scalar);
            ce.RegisterFunction("PRODUCT", 1, 255, Product, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("QUOTIENT", 2, 2, Adapt(Quotient), FunctionFlags.Scalar);
            ce.RegisterFunction("RADIANS", 1, 1, Adapt(Radians), FunctionFlags.Scalar);
            ce.RegisterFunction("RAND", 0, 0, Adapt(Rand), FunctionFlags.Scalar | FunctionFlags.Volatile);
            ce.RegisterFunction("RANDBETWEEN", 2, 2, Adapt(RandBetween), FunctionFlags.Scalar | FunctionFlags.Volatile);
            ce.RegisterFunction("ROMAN", 1, 2, AdaptLastOptional(Roman, 0), FunctionFlags.Scalar);
            ce.RegisterFunction("ROUND", 2, 2, Adapt(Round), FunctionFlags.Scalar);
            ce.RegisterFunction("ROUNDDOWN", 2, 2, Adapt(RoundDown), FunctionFlags.Scalar);
            ce.RegisterFunction("ROUNDUP", 2, 2, Adapt(RoundUp), FunctionFlags.Scalar);
            ce.RegisterFunction("SEC", 1, 1, Adapt(Sec), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("SECH", 1, 1, Adapt(Sech), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("SERIESSUM", 4, 4, AdaptSeriesSum(SeriesSum), FunctionFlags.Range, AllowRange.Only, 3);
            ce.RegisterFunction("SIGN", 1, 1, Adapt(Sign), FunctionFlags.Scalar);
            ce.RegisterFunction("SIN", 1, 1, Adapt(Sin), FunctionFlags.Scalar);
            ce.RegisterFunction("SINH", 1, 1, Adapt(Sinh), FunctionFlags.Scalar);
            ce.RegisterFunction("SQRT", 1, 1, Adapt(Sqrt), FunctionFlags.Scalar);
            ce.RegisterFunction("SQRTPI", 1, 1, Adapt(SqrtPi), FunctionFlags.Scalar);
            ce.RegisterFunction("SUBTOTAL", 2, 255, Adapt(Subtotal), FunctionFlags.Range, AllowRange.Except, 0);
            ce.RegisterFunction("SUM", 1, int.MaxValue, Sum, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("SUMIF", 2, 3, AdaptLastOptional(SumIf), FunctionFlags.Range, AllowRange.Only, 0, 2);
            ce.RegisterFunction("SUMIFS", 3, 255, AdaptIfs(SumIfs), FunctionFlags.Range, AllowRange.Only, new[] { 0 }.Concat(Enumerable.Range(0, 128).Select(x => x * 2 + 1)).ToArray());
            ce.RegisterFunction("SUMPRODUCT", 1, 30, Adapt(SumProduct), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("SUMSQ", 1, 255, SumSq, FunctionFlags.Range, AllowRange.All);
            //ce.RegisterFunction("SUMX2MY2", SumX2MY2, 1);
            //ce.RegisterFunction("SUMX2PY2", SumX2PY2, 1);
            //ce.RegisterFunction("SUMXMY2", SumXMY2, 1);
            ce.RegisterFunction("TAN", 1, 1, Adapt(Tan), FunctionFlags.Scalar);
            ce.RegisterFunction("TANH", 1, 1, Adapt(Tanh), FunctionFlags.Scalar);
            ce.RegisterFunction("TRUNC", 1, 2, AdaptLastOptional(Trunc, 0), FunctionFlags.Scalar);
        }

        #endregion Register

        public static double ASinh(double x)
        {
            return Math.Log(x + Math.Sqrt(x * x + 1.0));
        }

        public static double DegreesToGrads(double degrees)
        {
            return degrees / 9.0 * 10.0;
        }

        public static double DegreesToRadians(double degrees)
        {
            return Math.PI / 180.0 * degrees;
        }

        public static double GradsToDegrees(double grads)
        {
            return grads / 10.0 * 9.0;
        }

        public static double GradsToRadians(double grads)
        {
            return grads / 200.0 * Math.PI;
        }

        public static double RadiansToDegrees(double radians)
        {
            return 180.0 / Math.PI * radians;
        }

        public static double RadiansToGrads(double radians)
        {
            return radians / Math.PI * 200.0;
        }

        private static ScalarValue Abs(double number)
        {
            return Math.Abs(number);
        }

        private static ScalarValue Acos(double number)
        {
            if (Math.Abs(number) > 1)
                return XLError.NumberInvalid;

            return Math.Acos(number);
        }

        private static ScalarValue Acosh(double number)
        {
            if (number < 1)
                return XLError.NumberInvalid;

            return XLMath.ACosh(number);
        }

        private static ScalarValue Acot(double angle)
        {
            if (angle == 0)
                return Math.PI / 2;

            var acot = Math.Atan(1.0 / angle);

            // Acot in Excel calculates the modulus of the function above.
            // as the % operator is not the modulus, but the remainder, we have to calculate the modulus by hand:
            while (acot < 0)
                acot += Math.PI;

            return acot;
        }

        private static ScalarValue Acoth(double angle)
        {
            if (Math.Abs(angle) < 1)
                return XLError.NumberInvalid;

            return 0.5 * Math.Log((angle + 1) / (angle - 1));
        }

        private static ScalarValue Arabic(string input)
        {
            if (input.Length > 255)
                return XLError.IncompatibleValue;

            // Check minus sign
            var text = input.AsSpan().Trim();
            var minusSign = text.Length > 0 && text[0] == '-';
            if (minusSign)
                text = text[1..];

            var total = 0;
            for (var i = text.Length - 1; i >= 0; --i)
            {
                var addSymbol = char.ToUpperInvariant(text[i]);
                if (!RomanSymbolValues.TryGetValue(addSymbol, out var addValue))
                    return XLError.IncompatibleValue;

                total += addValue;

                // Standard roman numbers allow only one subtract symbol, Excel allows many
                // subtract symbols of different types.
                while (i > 0)
                {
                    var subtractSymbol = char.ToUpperInvariant(text[i - 1]);
                    if (!RomanSymbolValues.TryGetValue(subtractSymbol, out var subtractValue))
                        return XLError.IncompatibleValue;

                    if (subtractValue >= addValue)
                        break;

                    total -= subtractValue;
                    --i;
                }
            }

            if (minusSign && total == 0)
                return XLError.NumberInvalid;

            return minusSign ? -total : total;
        }

        private static ScalarValue Asin(double number)
        {
            if (Math.Abs(number) > 1)
                return XLError.NumberInvalid;

            return Math.Asin(number);
        }

        private static ScalarValue Asinh(double number)
        {
            return XLMath.Asinh(number);
        }

        private static ScalarValue Atan(double number)
        {
            return Math.Atan(number);
        }

        private static ScalarValue Atan2(double x, double y)
        {
            if (x == 0 && y == 0)
                return XLError.DivisionByZero;

            return Math.Atan2(y, x);
        }

        private static ScalarValue Atanh(double number)
        {
            if (Math.Abs(number) >= 1)
                return XLError.NumberInvalid;

            return XLMath.ATanh(number);
        }

        private static ScalarValue Base(double number, double radix, double minLength)
        {
            number = Math.Truncate(number);
            radix = Math.Truncate(radix);
            minLength = Math.Truncate(minLength);
            if (number is < 0 or > MaxDoubleInt || radix is < 2 or > 36 || minLength is < 0 or > 255)
                return XLError.NumberInvalid;

            var sb = new StringBuilder();
            while (number > 0)
            {
                var digit = (int)(number % radix);
                number = Math.Floor(number / radix);

                var digitChar = digit < 10
                    ? (char)(digit + '0')
                    : (char)(digit - 10 + 'A');
                sb.Insert(0, digitChar);
            }

            return sb.ToString().PadLeft((int)minLength, '0');
        }

        private static ScalarValue Ceiling(double number, double significance)
        {
            if (significance == 0)
                return 0;

            if (significance < 0 && number > 0)
                return XLError.NumberInvalid;

            if (number < 0)
                return -Math.Ceiling(-number / -significance) * -significance;

            return Math.Ceiling(number / significance) * significance;
        }

        private static ScalarValue CeilingMath(double number, double significance, double mode)
        {
            if (significance == 0)
                return 0;

            significance = Math.Abs(significance);

            // Mode 1 very similar to behavior of CEILING function, i.e. ceil
            // away from zero even for negative numbers. Mode 1 is not the same
            // as CEILING, e.g. CEILING(-5.5, 2.1) vs CEILING.MATH(-5.5, 2.1, 1)).
            if (number < 0 && mode != 0)
                return Math.Floor(number / significance) * significance;

            return Math.Ceiling(number / significance) * significance;
        }

        private static ScalarValue Combin(double number, double numberChosen)
        {
            var combinationsResult = XLMath.CombinChecked(number, numberChosen);
            if (!combinationsResult.TryPickT0(out var combinations, out var error))
                return error;

            return combinations;
        }

        private static ScalarValue CombinA(double number, double chosen)
        {
            number = Math.Truncate(number); // casting truncates towards 0 as specified
            chosen = Math.Truncate(chosen);

            if (number < 0)
                return XLError.NumberInvalid;

            if (chosen < 0)
                return XLError.NumberInvalid;

            var n = number + chosen - 1;
            if (n > int.MaxValue)
                return XLError.NumberInvalid;

            var k = number - 1;
            return chosen == 0 || k == 0
                ? 1
                : XLMath.Combin(n, k);
        }

        private static ScalarValue Cos(double number)
        {
            return Math.Cos(number);
        }

        private static ScalarValue Cosh(double number)
        {
            var cosh = Math.Cosh(number);
            if (double.IsInfinity(cosh))
                return XLError.NumberInvalid;

            return cosh;
        }

        private static ScalarValue Cot(double angle)
        {
            var tan = Math.Tan(angle);
            if (tan == 0)
                return XLError.DivisionByZero;

            return 1 / tan;
        }

        private static ScalarValue Coth(double angle)
        {
            if (angle == 0)
                return XLError.DivisionByZero;

            return 1 / Math.Tanh(angle);
        }

        private static ScalarValue Csc(double angle)
        {
            if (angle == 0)
                return XLError.DivisionByZero;

            return 1 / Math.Sin(angle);
        }

        private static ScalarValue Csch(double angle)
        {
            if (angle == 0)
                return XLError.DivisionByZero;

            return 1 / Math.Sinh(angle);
        }

        private static ScalarValue Decimal(string text, double radix)
        {
            radix = Math.Truncate(radix);
            if (radix is < 2 or > 36)
                return XLError.NumberInvalid;

            if (text.Length > 255)
                return XLError.IncompatibleValue;

            var result = 0d;
            foreach (var digit in text.AsSpan().TrimStart())
            {
                var digitNumber = digit switch
                {
                    >= '0' and <= '9' => digit - '0',
                    >= 'A' and <= 'Z' => digit - 'A' + 10,
                    >= 'a' and <= 'z' => digit - 'a' + 10,
                    _ => int.MaxValue,
                };

                if (digitNumber > radix - 1)
                    return XLError.NumberInvalid;

                result = result * radix + digitNumber;

                if (double.IsInfinity(result))
                    return XLError.NumberInvalid;
            }

            return result;
        }

        private static ScalarValue Degrees(double number)
        {
            return number * (180.0 / Math.PI);
        }

        private static ScalarValue Even(double number)
        {
            var num = Math.Ceiling(number);
            var addValue = num >= 0 ? 1 : -1;
            return XLMath.IsEven(num) ? num : num + addValue;
        }

        private static ScalarValue Exp(double number)
        {
            var exp = Math.Exp(number);
            if (double.IsInfinity(exp))
                return XLError.NumberInvalid;

            return exp;
        }

        private static ScalarValue Fact(double n)
        {
            if (n is < 0 or >= 171)
                return XLError.NumberInvalid;

            return XLMath.Factorial((int)Math.Floor(n));
        }

        private static ScalarValue FactDouble(double n)
        {
            var num = Math.Floor(n);
            if (num < -1)
                return XLError.NumberInvalid;

            var fact = 1.0;

            if (num > 1)
            {
                var start = XLMath.IsEven(num) ? 2 : 1;
                for (var i = start; i <= num; i += 2)
                {
                    fact *= i;
                    if (double.IsInfinity(fact))
                        return XLError.NumberInvalid;
                }
            }

            return fact;
        }

        private static ScalarValue Floor(double number, double significance)
        {
            // Rounding down, to zero. If we are at the zero, there is nowhere to go.
            if (number == 0)
                return 0;

            if (number > 0 && significance < 0)
                return XLError.NumberInvalid;

            if (significance == 0)
                return XLError.DivisionByZero;

            if (significance < 0)
                return -Math.Floor(-number / -significance) * -significance;

            return Math.Floor(number / significance) * significance;
        }

        private static ScalarValue FloorMath(double number, double significance, double mode)
        {
            if (significance == 0)
                return 0d;

            significance = Math.Abs(significance);
            if (number >= 0)
                return Math.Floor(number / significance) * significance;

            // Mode 0 floors numbers to lower number.
            if (mode == 0)
                return Math.Floor(number / significance) * significance;

            // Mode !0 truncates negative number, i.e. closer to zero
            return Math.Truncate(number / significance) * significance;
        }

        private static ScalarValue Gcd(CalcContext ctx, List<Array> arrays)
        {
            var result = 0d;
            foreach (var array in arrays)
            {
                foreach (var scalar in array)
                {
                    ctx.ThrowIfCancelled();
                    if (scalar.IsLogical)
                        return XLError.IncompatibleValue;

                    if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                        return error;

                    if (number is < 0 or > MaxDoubleInt)
                        return XLError.NumberInvalid;

                    result = Gcd(number, Math.Truncate(result));
                }
            }

            return result;
        }

        private static double Gcd(double a, double b)
        {
            a = Math.Truncate(a);
            b = Math.Truncate(b);
            while (b != 0)
                (a, b) = (b, a % b);

            return a;
        }

        private static double[,] GetArray(Expression expression)
        {
            if (expression is XObjectExpression objectExpression
                && objectExpression.Value is CellRangeReference cellRangeReference)
            {
                var range = cellRangeReference.Range;
                var rowCount = range.RowCount();
                var columnCount = range.ColumnCount();
                var arr = new double[rowCount, columnCount];

                for (int row = 0; row < rowCount; row++)
                {
                    for (int column = 0; column < columnCount; column++)
                    {
                        arr[row, column] = range.Cell(row + 1, column + 1).GetDouble();
                    }
                }

                return arr;
            }
            else
            {
                return new[,] { { (double)expression } };
            }
        }

        private static OneOf<double[,], XLError> GetArray(AnyValue value, CalcContext ctx)
        {
            if (value.TryPickSingleOrMultiValue(out var scalar, out var array, ctx))
                array = new ScalarArray(scalar, 1, 1);

            var rows = array.Height;
            var cols = array.Width;
            var arr = new double[rows, cols];

            for (var row = 0; row < rows; row++)
            {
                for (var col = 0; col < cols; col++)
                {
                    if (!array[row, col].TryPickNumber(out var number, out var error))
                        return error;

                    arr[row, col] = number;
                }
            }

            return arr;
        }

        private static ScalarValue Int(double number)
        {
            return Math.Floor(number);
        }

        private static ScalarValue Lcm(CalcContext ctx, List<Array> arrays)
        {
            var result = 1d;
            foreach (var array in arrays)
            {
                foreach (var scalar in array)
                {
                    ctx.ThrowIfCancelled();
                    if (scalar.IsLogical)
                        return XLError.IncompatibleValue;

                    if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                        return error;

                    if (number is < 0 or > MaxDoubleInt)
                        return XLError.NumberInvalid;

                    result = Lcm(result, Math.Truncate(number));
                }
            }

            return result;
        }

        private static double Lcm(double a, double b)
        {
            if (a == 0 || b == 0) return 0;
            return a * (b / Gcd(a, b));
        }

        private static ScalarValue Ln(double x)
        {
            if (x <= 0)
                return XLError.NumberInvalid;

            return Math.Log(x);
        }

        private static ScalarValue Log(double x, double @base)
        {
            if (x <= 0 || @base <= 0)
                return XLError.NumberInvalid;

            if (Math.Abs(@base - 1.0) < XLHelper.Epsilon)
                return XLError.DivisionByZero;

            return Math.Log(x, @base);
        }

        private static ScalarValue Log10(double x)
        {
            if (x <= 0)
                return XLError.NumberInvalid;

            return Math.Log10(x);
        }

        private static AnyValue MDeterm(CalcContext ctx, AnyValue value)
        {
            if (!GetArray(value, ctx).TryPickT0(out var array, out var error))
                return error;

            var isSquare = array.GetLength(0) == array.GetLength(1);
            if (!isSquare)
                return XLError.IncompatibleValue;

            var matrix = new XLMatrix(array);
            return matrix.Determinant();
        }

        private static AnyValue MInverse(CalcContext ctx, AnyValue value)
        {
            if (!GetArray(value, ctx).TryPickT0(out var array, out var error))
                return error;

            var isSquare = array.GetLength(0) == array.GetLength(1);
            if (!isSquare)
                return XLError.IncompatibleValue;

            var matrix = new XLMatrix(array);
            var inverse = matrix.Invert();
            if (inverse.IsSingular())
                return XLError.NumberInvalid;

            return new NumberArray(inverse.mat);
        }

        private static object MMult(List<Expression> p)
        {
            Double[,] A, B;

            try
            {
                A = GetArray(p[0]);
                B = GetArray(p[1]);
            }
            catch (InvalidCastException)
            {
                return XLError.IncompatibleValue;
            }

            if (A.GetLength(1) != B.GetLength(0))
                return XLError.IncompatibleValue;

            var C = new double[A.GetLength(0), B.GetLength(1)];
            for (int i = 0; i < A.GetLength(0); i++)
            {
                for (int j = 0; j < B.GetLength(1); j++)
                {
                    for (int k = 0; k < A.GetLength(1); k++)
                    {
                        C[i, j] += A[i, k] * B[k, j];
                    }
                }
            }

            return C;
        }

        private static ScalarValue Mod(double number, double divisor)
        {
            if (divisor == 0)
                return XLError.DivisionByZero;

            return number - Math.Floor(number / divisor) * divisor;
        }

        private static ScalarValue MRound(double number, double multiple)
        {
            if (multiple == 0)
                return 0;

            if (Math.Sign(number) != Math.Sign(multiple))
                return XLError.NumberInvalid;

            return Math.Round(number / multiple, MidpointRounding.AwayFromZero) * multiple;
        }

        private static ScalarValue Multinomial(CalcContext ctx, List<IEnumerable<ScalarValue>> numberCollections)
        {
            var numbersSum = 0.0;
            var denominator = 1.0;
            foreach (var numberCollection in numberCollections)
            {
                foreach (var scalar in numberCollection)
                {
                    ctx.ThrowIfCancelled();
                    if (scalar.IsLogical)
                        return XLError.IncompatibleValue;

                    if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                        return error;

                    if (number < 0)
                        return XLError.NumberInvalid;

                    number = Math.Truncate(number);
                    numbersSum += number;
                    denominator *= XLMath.Factorial(number);
                    if (double.IsInfinity(denominator))
                        return XLError.NumberInvalid;
                }
            }

            var numerator = XLMath.Factorial(numbersSum);
            if (double.IsInfinity(numerator))
                return XLError.NumberInvalid;

            return numerator / denominator;
        }

        private static ScalarValue Odd(double number)
        {
            var num = Math.Ceiling(number);
            var addValue = num >= 0 ? 1 : -1;
            return XLMath.IsOdd(num) ? num : num + addValue;
        }

        private static ScalarValue Pi()
        {
            return Math.PI;
        }

        private static ScalarValue Power(double x, double y)
        {
            // The value of x is negative and y is not a whole number, #NUM! is returned.
            var isPowerFraction = y % 1 != 0;
            if (x < 0 && isPowerFraction)
                return XLError.NumberInvalid;

            if (x == 0 && y == 0)
                return XLError.NumberInvalid;

            if (x == 0 && y < 0)
                return XLError.DivisionByZero;

            var power = Math.Pow(x, y);
            if (double.IsInfinity(power) || double.IsNaN(power))
                return XLError.NumberInvalid;

            return power;
        }

        private static AnyValue Product(CalcContext ctx, Span<AnyValue> args)
        {
            return Product(ctx, args, TallyNumbers.WithoutScalarBlank);
        }

        private static AnyValue Product(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            var result = tally.Tally(ctx, args, new ProductState(1, false));
            if (!result.TryPickT0(out var state, out var error))
                return error;

            return state.HasValues ? state.Product : 0;
        }

        private static ScalarValue Quotient(double dividend, double divisor)
        {
            if (divisor == 0)
                return XLError.DivisionByZero;

            return Math.Truncate(dividend / divisor);
        }

        private static ScalarValue Radians(double angle)
        {
            return angle * Math.PI / 180.0;
        }

        private static ScalarValue Rand()
        {
            return _rnd.NextDouble();
        }

        private static ScalarValue RandBetween(double lowerBound, double upperBound)
        {
            if (lowerBound > upperBound)
                return XLError.NumberInvalid;

            lowerBound = Math.Ceiling(lowerBound);
            upperBound = Math.Ceiling(upperBound);

            var range = upperBound - lowerBound;
            return lowerBound + Math.Round(_rnd.NextDouble() * range, MidpointRounding.AwayFromZero);
        }

        private static ScalarValue Roman(double number, double formValue)
        {
            if (number == 0)
                return string.Empty;

            if (number is < 0 or > 3999)
                return XLError.IncompatibleValue;

            var form = (int)Math.Truncate(formValue);
            if (form is < 0 or > 4)
                return XLError.IncompatibleValue;

            // The result can have at most 15 chars
            var result = new StringBuilder(15);
            var subtractValues = RomanForms.Value[form];
            foreach (var subtract in subtractValues)
            {
                // While the number is larger than the current value, append the symbol
                while (number >= subtract.Value)
                {
                    result.Append(subtract.Symbol);
                    number -= subtract.Value;
                }
            }

            return result.ToString();
        }

        private static ScalarValue Round(double value, double digits)
        {
            var digitCount = (int)Math.Truncate(digits);
            if (digits < 0)
            {
                var coef = Math.Pow(10, Math.Abs(digits));
                var shifted = value / coef;
                shifted = Math.Round(shifted, 0, MidpointRounding.AwayFromZero);
                return shifted * coef;
            }

            return Math.Round(value, digitCount, MidpointRounding.AwayFromZero);
        }

        private static ScalarValue RoundDown(double value, double digits)
        {
            var coef = Math.Pow(10, Math.Truncate(digits));
            return Math.Truncate(value * coef) / coef;
        }

        private static ScalarValue RoundUp(double value, double digits)
        {
            var coef = Math.Pow(10, Math.Truncate(digits));
            if (value >= 0)
                return Math.Ceiling(value * coef) / coef;

            return Math.Floor(value * coef) / coef;
        }

        private static ScalarValue Sec(double angle)
        {
            // Cos is actually never 0, because PI/2 can't be represented
            // as a double. It's just a really small number and the result
            // is thus never infinity.
            return 1.0 / Math.Cos(angle);
        }

        private static ScalarValue Sech(double angle)
        {
            return 1.0 / Math.Cosh(angle);
        }

        private static ScalarValue SeriesSum(CalcContext ctx, double input, double initial, double step, Array coefficients)
        {
            var total = 0d;
            var i = 0;
            foreach (var coefScalar in coefficients)
            {
                ctx.ThrowIfCancelled();
                if (!coefScalar.TryPickNumberOrBlank(out var coef, out var error))
                    return error;

                total += coef * Math.Pow(input, initial + i * step);
                if (double.IsInfinity(total))
                    return XLError.NumberInvalid;

                i++;
            }

            return total;
        }

        private static ScalarValue Sign(double number)
        {
            return Math.Sign(number);
        }

        private static ScalarValue Sin(double radians)
        {
            return Math.Sin(radians);
        }

        private static ScalarValue Sinh(double number)
        {
            var sinh = Math.Sinh(number);
            if (double.IsInfinity(sinh))
                return XLError.NumberInvalid;

            return sinh;
        }

        private static ScalarValue Sqrt(double number)
        {
            if (number < 0)
                return XLError.NumberInvalid;

            return Math.Sqrt(number);
        }

        private static ScalarValue SqrtPi(double number)
        {
            if (number < 0)
                return XLError.NumberInvalid;

            return Math.Sqrt(Math.PI * number);
        }

        private static AnyValue Subtotal(CalcContext ctx, double number, AnyValue[] fnArgs)
        {
            var funcNumber = number switch
            {
                >= 1 and < 12 => (int)number,
                >= 101 and < 112 => (int)number,
                _ => -1,
            };

            if (funcNumber < 0)
                return XLError.IncompatibleValue;

            var args = fnArgs.AsSpan();
            return funcNumber switch
            {
                1 => Statistical.Average(ctx, args, TallyNumbers.Subtotal10),
                2 => Statistical.Count(ctx, args, TallyNumbers.Subtotal10),
                3 => Statistical.Count(ctx, args, TallyAll.Subtotal10),
                4 => Statistical.Max(ctx, args, TallyNumbers.Subtotal10),
                5 => Statistical.Min(ctx, args, TallyNumbers.Subtotal10),
                6 => Product(ctx, args, TallyNumbers.Subtotal10),
                7 => Statistical.StDev(ctx, args, TallyNumbers.Subtotal10),
                8 => Statistical.StDevP(ctx, args, TallyNumbers.Subtotal10),
                9 => Sum(ctx, args, TallyNumbers.Subtotal10),
                10 => Statistical.Var(ctx, args, TallyNumbers.Subtotal10),
                11 => Statistical.VarP(ctx, args, TallyNumbers.Subtotal10),
                101 => Statistical.Average(ctx, args, TallyNumbers.Subtotal100),
                102 => Statistical.Count(ctx, args, TallyNumbers.Subtotal100),
                103 => Statistical.Count(ctx, args, TallyAll.Subtotal100),
                104 => Statistical.Max(ctx, args, TallyNumbers.Subtotal100),
                105 => Statistical.Min(ctx, args, TallyNumbers.Subtotal100),
                106 => Product(ctx, args, TallyNumbers.Subtotal100),
                107 => Statistical.StDev(ctx, args, TallyNumbers.Subtotal100),
                108 => Statistical.StDevP(ctx, args, TallyNumbers.Subtotal100),
                109 => Sum(ctx, args, TallyNumbers.Subtotal100),
                110 => Statistical.Var(ctx, args, TallyNumbers.Subtotal100),
                111 => Statistical.VarP(ctx, args, TallyNumbers.Subtotal100),
                _ => throw new UnreachableException(),
            };
        }

        private static AnyValue Sum(CalcContext ctx, Span<AnyValue> args)
        {
            return Sum(ctx, args, TallyNumbers.Default);
        }

        private static AnyValue Sum(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            var result = tally.Tally(ctx, args, new SumState(0));
            if (!result.TryPickT0(out var state, out var error))
                return error;

            return state.Sum;
        }

        private static AnyValue SumIf(CalcContext ctx, AnyValue range, ScalarValue selectionCriteria, AnyValue sumRange)
        {
            // Sum range is optional. If not specified, use the range as the sum range.
            if (sumRange.IsBlank)
                sumRange = range;

            var tally = new TallyCriteria();
            var criteria = Criteria.Create(selectionCriteria, ctx.Culture);

            // Excel doesn't support anything but area in the syntax, but we need to deal with it somehow.
            if (!range.TryPickArea(out var area, out var areaError))
                return areaError;

            if (!sumRange.TryPickArea(out _, out var sumAreaError))
                return sumAreaError;

            tally.Add(area, criteria);

            return Sum(ctx, new[] { sumRange }, tally);
        }

        private static AnyValue SumIfs(CalcContext ctx, AnyValue sumRange, List<(AnyValue Range, ScalarValue Criteria)> criteriaRanges)
        {
            if (!sumRange.TryPickArea(out var sumArea, out var sumAreaError))
                return sumAreaError;

            var tally = new TallyCriteria();
            foreach (var (selectionRange, selectionCriteria) in criteriaRanges)
            {
                var criteria = Criteria.Create(selectionCriteria, ctx.Culture);
                if (!selectionRange.TryPickArea(out var selectionArea, out var selectionAreaError))
                    return selectionAreaError;

                // All areas must have same size, that is different
                // from SUMIF where areas can have different size.
                if (sumArea.RowSpan != selectionArea.RowSpan ||
                    sumArea.ColumnSpan != selectionArea.ColumnSpan)
                    return XLError.IncompatibleValue;

                tally.Add(selectionArea, criteria);
            }

            return Sum(ctx, new[] { sumRange }, tally);
        }

        private static AnyValue SumProduct(CalcContext _, Array[] areas)
        {
            if (areas.Length < 1)
                return XLError.IncompatibleValue;

            var width = 0;
            var height = 0;

            // Check that all arguments have same width and height.
            foreach (var area in areas)
            {
                var areaWidth = area.Width;
                var areaHeight = area.Height;

                // We don't need to do this check for every value later, because scalar
                // blank value can only happen for 1x1.
                if (areaWidth == 1 &&
                    areaHeight == 1 &&
                    area[0, 0].IsBlank)
                    return XLError.IncompatibleValue;

                // If this is the first argument, use it as a baseline width and height
                if (width == 0) width = areaWidth;
                if (height == 0) height = areaHeight;

                if (width != areaWidth || height != areaHeight)
                    return XLError.IncompatibleValue;
            }

            // Calculate SumProduct
            var sum = 0.0;
            for (var rowIdx = 0; rowIdx < height; ++rowIdx)
            {
                for (var colIdx = 0; colIdx < width; ++colIdx)
                {
                    var product = 1.0;
                    foreach (var area in areas)
                    {
                        var scalar = area[rowIdx, colIdx];

                        if (scalar.TryPickError(out var error))
                            return error;

                        if (!scalar.TryPickNumber(out var number))
                            number = 0;

                        product *= number;
                    }

                    sum += product;
                }
            }

            return sum;
        }

        private static AnyValue SumSq(CalcContext ctx, Span<AnyValue> args)
        {
            var result = TallyNumbers.Default.Tally(ctx, args, new SumSqState(0.0));
            if (!result.TryPickT0(out var sumSq, out var error))
                return error;

            return sumSq.Sum;
        }

        private static ScalarValue Tan(double radians)
        {
            // Cutoff point for Excel. .NET Core allows all values and .NET Fx ~< 1e+19.
            // To ensure consistent behavior for all platforms, respect Excel limit. It's
            // lower than both the .NET Core and the .NET Fx one.
            if (Math.Abs(radians) >= 134217728)
                return XLError.NumberInvalid;

            return Math.Tan(radians);
        }

        private static ScalarValue Tanh(double number)
        {
            return Math.Tanh(number);
        }

        private static ScalarValue Trunc(double number, double digits)
        {
            var scaling = Math.Pow(10, digits);
            return Math.Truncate(number * scaling) / scaling;
        }

        private static Dictionary<int, IReadOnlyList<(string Symbol, int Value)>> BuildRomanForms()
        {
            // Roman numbers can have several forms and each one has a different set of possible values.
            // In Excel, each successive one has more subtract values than previous one.
            var allForms = new Dictionary<int, IReadOnlyList<(string Symbol, int Value)>>();
            var form0 = new List<(string Symbol, int Value)>
            {
                ("M", 1000), ("CM", 900),
                ("D", 500), ("CD", 400),
                ("C", 100), ("XC", 90),
                ("L", 50), ("XL", 40),
                ("X", 10), ("IX", 9),
                ("V", 5), ("IV", 4),
                ("I", 1),
            };
            allForms.Add(0, form0);

            var form1Additions = new (string Symbol, int Value)[]
            {
                ("LM", 950),
                ("LD", 450),
                ("VC", 95),
                ("VL", 45),
            };
            var form1 = form0.Concat(form1Additions).OrderByDescending(x => x.Value).ToArray();
            allForms.Add(1, form1);

            var form2Additions = new (string Symbol, int Value)[]
            {
                ("XM", 990),
                ("XD", 490),
                ("IC", 99),
                ("IL", 49),
            };
            var form2 = form1.Concat(form2Additions).OrderByDescending(x => x.Value).ToArray();
            allForms.Add(2, form2);

            var form3Additions = new (string Symbol, int Value)[]
            {
                ("VM", 995),
                ("VD", 495),
            };
            var form3 = form2.Concat(form3Additions).OrderByDescending(x => x.Value).ToArray();
            allForms.Add(3, form3);

            var form4Additions = new (string Symbol, int Value)[]
            {
                ("IM", 999),
                ("ID", 499),
            };
            var form4 = form3.Concat(form4Additions).OrderByDescending(x => x.Value).ToArray();
            allForms.Add(4, form4);
            return allForms;
        }

        private readonly record struct SumState(double Sum) : ITallyState<SumState>
        {
            public SumState Tally(double number) => new(Sum + number);
        }

        private readonly record struct SumSqState(double Sum) : ITallyState<SumSqState>
        {
            public SumSqState Tally(double number)
            {
                return new SumSqState(Sum + number * number);
            }
        }

        private readonly record struct ProductState(double Product, bool HasValues) : ITallyState<ProductState>
        {
            public ProductState Tally(double number) => new(Product * number, true);
        }
    }
}
