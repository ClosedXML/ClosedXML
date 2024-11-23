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
        private static readonly Random _rnd = new Random();

        /// <summary>
        /// Key: roman form. Value: A collection of subtract symbols and subtract value.
        /// Collection is sorted by subtract value in descending order.
        /// </summary>
        private static readonly Lazy<IReadOnlyDictionary<int, IReadOnlyList<(string Symbol, int Value)>>> RomanForms = new(BuildRomanForms);

        #region Register

        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("ABS", 1, 1, Adapt(Abs), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOS", 1, 1, Adapt(Acos), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOSH", 1, 1, Adapt(Acosh), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOT", 1, Acot);
            ce.RegisterFunction("ACOTH", 1, Acoth);
            ce.RegisterFunction("ARABIC", 1, Arabic);
            ce.RegisterFunction("ASIN", 1, 1, Adapt(Asin), FunctionFlags.Scalar);
            ce.RegisterFunction("ASINH", 1, 1, Adapt(Asinh), FunctionFlags.Scalar);
            ce.RegisterFunction("ATAN", 1, 1, Adapt(Atan), FunctionFlags.Scalar);
            ce.RegisterFunction("ATAN2", 2, 2, Adapt(Atan2), FunctionFlags.Scalar);
            ce.RegisterFunction("ATANH", 1, 1, Adapt(Atanh), FunctionFlags.Scalar);
            ce.RegisterFunction("BASE", 2, 3, Base);
            ce.RegisterFunction("CEILING", 2, 2, Adapt(Ceiling), FunctionFlags.Scalar);
            ce.RegisterFunction("CEILING.MATH", 1, 3, AdaptLastTwoOptional(CeilingMath, 1, 0), FunctionFlags.Scalar | FunctionFlags.Future);
            ce.RegisterFunction("COMBIN", 2, 2, Adapt(Combin), FunctionFlags.Scalar);
            ce.RegisterFunction("COMBINA", 2, CombinA);
            ce.RegisterFunction("COS", 1, 1, Adapt(Cos), FunctionFlags.Scalar);
            ce.RegisterFunction("COSH", 1, 1, Adapt(Cosh), FunctionFlags.Scalar);
            ce.RegisterFunction("COT", 1, Cot);
            ce.RegisterFunction("COTH", 1, Coth);
            ce.RegisterFunction("CSC", 1, Csc);
            ce.RegisterFunction("CSCH", 1, Csch);
            ce.RegisterFunction("DECIMAL", 2, MathTrig.Decimal);
            ce.RegisterFunction("DEGREES", 1, 1, Adapt(Degrees), FunctionFlags.Scalar);
            ce.RegisterFunction("EVEN", 1, 1, Adapt(Even), FunctionFlags.Scalar);
            ce.RegisterFunction("EXP", 1, 1, Adapt(Exp), FunctionFlags.Scalar);
            ce.RegisterFunction("FACT", 1, 1, Adapt(Fact), FunctionFlags.Scalar);
            ce.RegisterFunction("FACTDOUBLE", 1, 1, Adapt(FactDouble), FunctionFlags.Scalar);
            ce.RegisterFunction("FLOOR", 2, 2, Adapt(Floor), FunctionFlags.Scalar);
            ce.RegisterFunction("FLOOR.MATH", 1, 3, FloorMath);
            ce.RegisterFunction("GCD", 1, 255, Gcd);
            ce.RegisterFunction("INT", 1, 1, Adapt(Int), FunctionFlags.Scalar);
            ce.RegisterFunction("LCM", 1, 255, Lcm);
            ce.RegisterFunction("LN", 1, 1, Adapt(Ln), FunctionFlags.Scalar);
            ce.RegisterFunction("LOG", 1, 2, AdaptLastOptional(Log, 10), FunctionFlags.Scalar);
            ce.RegisterFunction("LOG10", 1, 1, Adapt(Log10), FunctionFlags.Scalar);
            ce.RegisterFunction("MDETERM", 1, MDeterm, AllowRange.All);
            ce.RegisterFunction("MINVERSE", 1, MInverse, AllowRange.All);
            ce.RegisterFunction("MMULT", 2, MMult, AllowRange.All);
            ce.RegisterFunction("MOD", 2, 2, Adapt(Mod), FunctionFlags.Scalar);
            ce.RegisterFunction("MROUND", 2, 2, Adapt(MRound), FunctionFlags.Scalar);
            ce.RegisterFunction("MULTINOMIAL", 1, 255, Multinomial);
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
            ce.RegisterFunction("SEC", 1, Sec);
            ce.RegisterFunction("SECH", 1, Sech);
            ce.RegisterFunction("SERIESSUM", 4, SeriesSum, AllowRange.Only, 3);
            ce.RegisterFunction("SIGN", 1, Sign);
            ce.RegisterFunction("SIN", 1, Sin);
            ce.RegisterFunction("SINH", 1, Sinh);
            ce.RegisterFunction("SQRT", 1, Sqrt);
            ce.RegisterFunction("SQRTPI", 1, SqrtPi);
            ce.RegisterFunction("SUBTOTAL", 2, 255, Adapt(Subtotal), FunctionFlags.Range, AllowRange.Except, 0);
            ce.RegisterFunction("SUM", 1, int.MaxValue, Sum, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("SUMIF", 2, 3, AdaptLastOptional(SumIf), FunctionFlags.Range, AllowRange.Only, 0, 2);
            ce.RegisterFunction("SUMIFS", 3, 255, AdaptIfs(SumIfs), FunctionFlags.Range, AllowRange.Only, new[] { 0 }.Concat(Enumerable.Range(0, 128).Select(x => x * 2 + 1)).ToArray());
            ce.RegisterFunction("SUMPRODUCT", 1, 30, Adapt(SumProduct), FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("SUMSQ", 1, 255, SumSq, FunctionFlags.Range, AllowRange.All);
            //ce.RegisterFunction("SUMX2MY2", SumX2MY2, 1);
            //ce.RegisterFunction("SUMX2PY2", SumX2PY2, 1);
            //ce.RegisterFunction("SUMXMY2", SumXMY2, 1);
            ce.RegisterFunction("TAN", 1, Tan);
            ce.RegisterFunction("TANH", 1, Tanh);
            ce.RegisterFunction("TRUNC", 1, 2, Trunc);
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

        private static object Acot(List<Expression> p)
        {
            double x = Math.Atan(1.0 / p[0]);

            // Acot in Excel calculates the modulus of the function above.
            // as the % operator is not the modulus, but the remainder, we have to calculate the modulus by hand:
            while (x < 0)
                x += Math.PI;

            return x;
        }

        private static object Acoth(List<Expression> p)
        {
            double number = p[0];
            if (Math.Abs(number) < 1)
                return XLError.NumberInvalid;

            return 0.5 * Math.Log((number + 1) / (number - 1));
        }

        private static object Arabic(List<Expression> p)
        {
            string input = ((string)p[0]).Trim();

            try
            {
                if (input.Length == 0)
                    return 0;
                if (input == "-")
                    return XLError.NumberInvalid;
                else if (input[0] == '-')
                    return -XLMath.RomanToArabic(input.Substring(1));
                else
                    return XLMath.RomanToArabic(input);
            }
            catch (ArgumentOutOfRangeException)
            {
                return XLError.IncompatibleValue;
            }
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

        private static object Base(List<Expression> p)
        {
            long number;
            int radix;
            int minLength = 0;

            var rawNumber = p[0].Evaluate();
            if (rawNumber is long || rawNumber is int || rawNumber is byte || rawNumber is double || rawNumber is float)
                number = Convert.ToInt64(rawNumber);
            else
                return XLError.IncompatibleValue;

            var rawRadix = p[1].Evaluate();
            if (rawRadix is long || rawRadix is int || rawRadix is byte || rawRadix is double || rawRadix is float)
                radix = Convert.ToInt32(rawRadix);
            else
                return XLError.IncompatibleValue;

            if (p.Count > 2)
            {
                var rawMinLength = p[2].Evaluate();
                if (rawMinLength is long || rawMinLength is int || rawMinLength is byte || rawMinLength is double || rawMinLength is float)
                    minLength = Convert.ToInt32(rawMinLength);
                else
                    return XLError.IncompatibleValue;
            }

            if (number < 0 || radix < 2 || radix > 36)
                return XLError.NumberInvalid;

            return XLMath.ChangeBase(number, radix).PadLeft(minLength, '0');
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
            // as CEILING, e.g. CEILING(5.5, -2.1) vs CEILING.MATH(5.5, -2.1, 1)).
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

        private static object CombinA(List<Expression> p)
        {
            Int32 number = (int)p[0]; // casting truncates towards 0 as specified
            Int32 chosen = (int)p[1];

            if (number < 0 || number < chosen)
                return XLError.NumberInvalid;
            if (chosen < 0)
                return XLError.NumberInvalid;

            int n = number + chosen - 1;
            int k = number - 1;

            return n == k || k == 0
                ? 1
                : (long)XLMath.Combin(n, k);
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

        private static object Cot(List<Expression> p)
        {
            var tan = Math.Tan(p[0]);

            if (tan == 0)
                return XLError.DivisionByZero;

            return 1 / tan;
        }

        private static object Coth(List<Expression> p)
        {
            double input = p[0];
            if (input == 0)
                return XLError.DivisionByZero;

            return 1 / Math.Tanh(input);
        }

        private static object Csc(List<Expression> p)
        {
            double input = p[0];
            if (input == 0)
                return XLError.DivisionByZero;

            return 1 / Math.Sin(input);
        }

        private static object Csch(List<Expression> p)
        {
            if (Math.Abs((double)p[0].Evaluate()) < Double.Epsilon)
                return XLError.DivisionByZero;

            return 1 / Math.Sinh(p[0]);
        }

        private static object Decimal(List<Expression> p)
        {
            string source = p[0];
            double radix = p[1];

            if (radix < 2 || radix > 36)
                return XLError.NumberInvalid;

            var asciiValues = Encoding.ASCII.GetBytes(source.ToUpperInvariant());

            double result = 0;
            int i = 0;

            foreach (byte digit in asciiValues)
            {
                if (digit > 90)
                {
                    return XLError.NumberInvalid;
                }

                int digitNumber = digit >= 48 && digit < 58
                    ? digit - 48
                    : digit - 55;

                if (digitNumber > radix - 1)
                    return XLError.NumberInvalid;

                result = result * radix + digitNumber;
                i++;
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

        private static object FloorMath(List<Expression> p)
        {
            double number = p[0];
            double significance = 1;
            if (p.Count > 1) significance = p[1];

            double mode = 0;
            if (p.Count > 2) mode = p[2];

            if (significance == 0)
                return 0d;
            else if (number >= 0)
                return Math.Floor(number / Math.Abs(significance)) * Math.Abs(significance);
            else if (mode == 0)
                return Math.Floor(number / Math.Abs(significance)) * Math.Abs(significance);
            else
                return -Math.Floor(-number / Math.Abs(significance)) * Math.Abs(significance);
        }

        private static object Gcd(List<Expression> p)
        {
            return p.Select(v => (int)v).Aggregate(Gcd);
        }

        private static int Gcd(int a, int b)
        {
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

        private static ScalarValue Int(double number)
        {
            return Math.Floor(number);
        }

        private static object Lcm(List<Expression> p)
        {
            return p.Select(v => (int)v).Aggregate(Lcm);
        }

        private static int Lcm(int a, int b)
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

        private static object MDeterm(List<Expression> p)
        {
            var arr = GetArray(p[0]);
            var m = new XLMatrix(arr);

            return m.Determinant();
        }

        private static object MInverse(List<Expression> p)
        {
            var arr = GetArray(p[0]);
            var m = new XLMatrix(arr);

            return m.Invert().mat;
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

        private static object Multinomial(List<Expression> p)
        {
            return Multinomial(p.ConvertAll(v => (double)v));
        }

        private static double Multinomial(List<double> numbers)
        {
            double numbersSum = 0;
            foreach (var number in numbers)
            {
                numbersSum += number;
            }

            double maxNumber = numbers.Max();
            var denomFactorPowers = new double[(uint)numbers.Max() + 1];
            foreach (var number in numbers)
            {
                for (int i = 2; i <= number; i++)
                {
                    denomFactorPowers[i]++;
                }
            }

            for (int i = 2; i < denomFactorPowers.Length; i++)
            {
                denomFactorPowers[i]--; // reduce with nominator
            }

            int currentFactor = 2;
            double currentPower = 1;
            double result = 1;
            for (double i = maxNumber + 1; i <= numbersSum; i++)
            {
                double tempDenom = 1;
                while (tempDenom < result && currentFactor < denomFactorPowers.Length)
                {
                    if (currentPower > denomFactorPowers[currentFactor])
                    {
                        currentFactor++;
                        currentPower = 1;
                    }
                    else
                    {
                        tempDenom *= currentFactor;
                        currentPower++;
                    }
                }
                result = result / tempDenom * i;
            }

            return result;
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

        private static object Sec(List<Expression> p)
        {
            if (double.TryParse(p[0], out double number))
                return 1.0 / Math.Cos(number);
            else
                return XLError.IncompatibleValue;
        }

        private static object Sech(List<Expression> p)
        {
            return 1.0 / Math.Cosh(p[0]);
        }

        private static object SeriesSum(List<Expression> p)
        {
            var x = (Double)p[0];
            var n = (Double)p[1];
            var m = (Double)p[2];
            if (p[3] is XObjectExpression obj)
            {
                Double total = 0;
                Int32 i = 0;
                foreach (var e in obj)
                {
                    total += (double)e * Math.Pow(x, n + i * m);
                    i++;
                }

                return total;
            }
            else
            {
                return p[3] * Math.Pow(x, n);
            }
        }

        private static object Sign(List<Expression> p)
        {
            return Math.Sign(p[0]);
        }

        private static object Sin(List<Expression> p)
        {
            return Math.Sin(p[0]);
        }

        private static object Sinh(List<Expression> p)
        {
            return Math.Sinh(p[0]);
        }

        private static object Sqrt(List<Expression> p)
        {
            return Math.Sqrt(p[0]);
        }

        private static object SqrtPi(List<Expression> p)
        {
            var num = (Double)p[0];
            return Math.Sqrt(Math.PI * num);
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

        private static object Tan(List<Expression> p)
        {
            return Math.Tan(p[0]);
        }

        private static object Tanh(List<Expression> p)
        {
            return Math.Tanh(p[0]);
        }

        private static object Trunc(List<Expression> p)
        {
            var number = (double)p[0];

            var num_digits = 0d;
            if (p.Count > 1)
                num_digits = (double)p[1];

            var scaling = Math.Pow(10, num_digits);

            var truncated = (int)(number * scaling);
            return (double)truncated / scaling;
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
