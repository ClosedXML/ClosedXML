// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using ClosedXML.Excel.CalcEngine.Exceptions;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class MathTrig
    {
        private static readonly Random _rnd = new Random();

        #region Register

        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("ABS", 1, 1, Adapt(Abs), FunctionFlags.Scalar);
            ce.RegisterFunction("ACOS", 1, Acos);
            ce.RegisterFunction("ACOSH", 1, Acosh);
            ce.RegisterFunction("ACOT", 1, Acot);
            ce.RegisterFunction("ACOTH", 1, Acoth);
            ce.RegisterFunction("ARABIC", 1, Arabic);
            ce.RegisterFunction("ASIN", 1, Asin);
            ce.RegisterFunction("ASINH", 1, Asinh);
            ce.RegisterFunction("ATAN", 1, Atan);
            ce.RegisterFunction("ATAN2", 2, Atan2);
            ce.RegisterFunction("ATANH", 1, Atanh);
            ce.RegisterFunction("BASE", 2, 3, Base);
            ce.RegisterFunction("CEILING", 2, Ceiling);
            ce.RegisterFunction("CEILING.MATH", 1, 3, CeilingMath);
            ce.RegisterFunction("COMBIN", 2, 2, Adapt(Combin), FunctionFlags.Scalar);
            ce.RegisterFunction("COMBINA", 2, CombinA);
            ce.RegisterFunction("COS", 1, Cos);
            ce.RegisterFunction("COSH", 1, Cosh);
            ce.RegisterFunction("COT", 1, Cot);
            ce.RegisterFunction("COTH", 1, Coth);
            ce.RegisterFunction("CSC", 1, Csc);
            ce.RegisterFunction("CSCH", 1, Csch);
            ce.RegisterFunction("DECIMAL", 2, MathTrig.Decimal);
            ce.RegisterFunction("DEGREES", 1, Degrees);
            ce.RegisterFunction("EVEN", 1, Even);
            ce.RegisterFunction("EXP", 1, Exp);
            ce.RegisterFunction("FACT", 1, 1, Adapt(Fact), FunctionFlags.Scalar);
            ce.RegisterFunction("FACTDOUBLE", 1, FactDouble);
            ce.RegisterFunction("FLOOR", 2, Floor);
            ce.RegisterFunction("FLOOR.MATH", 1, 3, FloorMath);
            ce.RegisterFunction("GCD", 1, 255, Gcd);
            ce.RegisterFunction("INT", 1, Int);
            ce.RegisterFunction("LCM", 1, 255, Lcm);
            ce.RegisterFunction("LN", 1, Ln);
            ce.RegisterFunction("LOG", 1, 2, Log);
            ce.RegisterFunction("LOG10", 1, Log10);
            ce.RegisterFunction("MDETERM", 1, MDeterm, AllowRange.All);
            ce.RegisterFunction("MINVERSE", 1, MInverse, AllowRange.All);
            ce.RegisterFunction("MMULT", 2, MMult, AllowRange.All);
            ce.RegisterFunction("MOD", 2, Mod);
            ce.RegisterFunction("MROUND", 2, MRound);
            ce.RegisterFunction("MULTINOMIAL", 1, 255, Multinomial);
            ce.RegisterFunction("ODD", 1, Odd);
            ce.RegisterFunction("PI", 0, Pi);
            ce.RegisterFunction("POWER", 2, Power);
            ce.RegisterFunction("PRODUCT", 1, 255, Product, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("QUOTIENT", 2, Quotient);
            ce.RegisterFunction("RADIANS", 1, Radians);
            ce.RegisterFunction("RAND", 0, Rand);
            ce.RegisterFunction("RANDBETWEEN", 2, RandBetween);
            ce.RegisterFunction("ROMAN", 1, 2, Roman);
            ce.RegisterFunction("ROUND", 2, Round);
            ce.RegisterFunction("ROUNDDOWN", 2, RoundDown);
            ce.RegisterFunction("ROUNDUP", 1, 2, RoundUp);
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
            ce.RegisterFunction("SUMIF", 2, 3, SumIf, AllowRange.Only, 0, 2);
            ce.RegisterFunction("SUMIFS", 3, 255, SumIfs, AllowRange.Only, new[] { 0 }.Concat(Enumerable.Range(0, 128).Select(x => x * 2 + 1)).ToArray());
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

        private static AnyValue Abs(double number)
        {
            return Math.Abs(number);
        }

        private static object Acos(List<Expression> p)
        {
            double input = p[0];
            if (Math.Abs(input) > 1)
                return XLError.NumberInvalid;

            return Math.Acos(p[0]);
        }

        private static object Acosh(List<Expression> p)
        {
            double number = p[0];
            if (number < 1)
                return XLError.NumberInvalid;

            return XLMath.ACosh(p[0]);
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

        private static object Asin(List<Expression> p)
        {
            double input = p[0];
            if (Math.Abs(input) > 1)
                return XLError.NumberInvalid;

            return Math.Asin(input);
        }

        private static object Asinh(List<Expression> p)
        {
            return XLMath.ASinh(p[0]);
        }

        private static object Atan(List<Expression> p)
        {
            return Math.Atan(p[0]);
        }

        private static object Atan2(List<Expression> p)
        {
            double x = p[0];
            double y = p[1];
            if (x == 0 && y == 0)
                return XLError.DivisionByZero;

            return Math.Atan2(y, x);
        }

        private static object Atanh(List<Expression> p)
        {
            double input = p[0];
            if (Math.Abs(input) >= 1)
                return XLError.NumberInvalid;

            return XLMath.ATanh(p[0]);
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

        private static object Ceiling(List<Expression> p)
        {
            double number = p[0];
            double significance = p[1];

            if (significance == 0)
                return 0d;
            else if (significance < 0 && number > 0)
                return XLError.NumberInvalid;
            else if (significance < 0)
                return -Math.Ceiling(-number / -significance) * -significance;
            else
                return Math.Ceiling(number / significance) * significance;
        }

        private static object CeilingMath(List<Expression> p)
        {
            double number = p[0];
            double significance = 1;
            if (p.Count > 1) significance = p[1];

            double mode = 0;
            if (p.Count > 2) mode = p[2];

            if (significance == 0)
                return 0d;
            else if (number >= 0)
                return Math.Ceiling(number / Math.Abs(significance)) * Math.Abs(significance);
            else if (mode == 0)
                return Math.Ceiling(number / Math.Abs(significance)) * Math.Abs(significance);
            else
                return -Math.Ceiling(-number / Math.Abs(significance)) * Math.Abs(significance);
        }

        private static AnyValue Combin(double number, double numberChosen)
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

        private static object Cos(List<Expression> p)
        {
            return Math.Cos(p[0]);
        }

        private static object Cosh(List<Expression> p)
        {
            return Math.Cosh(p[0]);
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

        private static object Degrees(List<Expression> p)
        {
            return p[0] * (180.0 / Math.PI);
        }

        private static object Even(List<Expression> p)
        {
            var num = (int)Math.Ceiling(p[0]);
            var addValue = num >= 0 ? 1 : -1;
            return XLMath.IsEven(num) ? num : num + addValue;
        }

        private static object Exp(List<Expression> p)
        {
            return Math.Exp(p[0]);
        }

        private static AnyValue Fact(double n)
        {
            if (n is < 0 or >= 171)
                return XLError.NumberInvalid;

            return XLMath.Factorial((int)Math.Floor(n));
        }

        private static object FactDouble(List<Expression> p)
        {
            var input = p[0].Evaluate();

            if (!(input is long || input is int || input is byte || input is double || input is float))
                return XLError.IncompatibleValue;

            var num = Math.Floor(p[0]);
            double fact = 1.0;

            if (num < -1)
                return XLError.NumberInvalid;

            if (num > 1)
            {
                var start = Math.Abs(num % 2) < XLHelper.Epsilon ? 2 : 1;
                for (int i = start; i <= num; i += 2)
                    fact *= i;
            }
            return fact;
        }

        private static object Floor(List<Expression> p)
        {
            double number = p[0];
            double significance = p[1];

            if (significance == 0)
                return XLError.DivisionByZero;
            else if (significance < 0 && number > 0)
                return XLError.NumberInvalid;
            else if (significance < 0)
                return -Math.Floor(-number / -significance) * -significance;
            else
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
            return b == 0 ? a : Gcd(b, a % b);
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

        private static object Int(List<Expression> p)
        {
            return Math.Floor(p[0]);
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

        private static object Ln(List<Expression> p)
        {
            return Math.Log(p[0]);
        }

        private static object Log(List<Expression> p)
        {
            var lbase = p.Count > 1 ? (double)p[1] : 10;
            return Math.Log(p[0], lbase);
        }

        private static object Log10(List<Expression> p)
        {
            return Math.Log10(p[0]);
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

        private static object Mod(List<Expression> p)
        {
            double number = p[0];
            double divisor = p[1];

            return number - Math.Floor(number / divisor) * divisor;
        }

        private static object MRound(List<Expression> p)
        {
            var number = (Double)p[0];
            var multiple = (Double)p[1];

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

        private static object Odd(List<Expression> p)
        {
            var num = (int)Math.Ceiling(p[0]);
            var addValue = num >= 0 ? 1 : -1;
            return XLMath.IsOdd(num) ? num : num + addValue;
        }

        private static object Pi(List<Expression> p)
        {
            return Math.PI;
        }

        private static object Power(List<Expression> p)
        {
            return Math.Pow(p[0], p[1]);
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

        private static object Quotient(List<Expression> p)
        {
            Double n = p[0];
            Double k = p[1];

            return (int)(n / k);
        }

        private static object Radians(List<Expression> p)
        {
            return p[0] * Math.PI / 180.0;
        }

        private static object Rand(List<Expression> p)
        {
            return _rnd.NextDouble();
        }

        private static object RandBetween(List<Expression> p)
        {
            return _rnd.Next((int)(double)p[0], (int)(double)p[1]);
        }

        private static object Roman(List<Expression> p)
        {
            if (p.Count == 1
                || (Boolean.TryParse((string)p[1], out bool boolTemp) && boolTemp)
                || (Int32.TryParse((string)p[1], out int intTemp) && intTemp == 1))
            {
                return XLMath.ToRoman((int)p[0]);
            }

            throw new ArgumentException("Can only support classic roman types.");
        }

        private static object Round(List<Expression> p)
        {
            var value = (Double)p[0];
            var digits = (Int32)(Double)p[1];
            if (digits >= 0)
            {
                return Math.Round(value, digits, MidpointRounding.AwayFromZero);
            }
            else
            {
                digits = Math.Abs(digits);
                double temp = value / Math.Pow(10, digits);
                temp = Math.Round(temp, 0, MidpointRounding.AwayFromZero);
                return temp * Math.Pow(10, digits);
            }
        }

        private static object RoundDown(List<Expression> p)
        {
            var value = (Double)p[0];
            var digits = (Int32)(Double)p[1];

            if (value >= 0)
                return Math.Floor(value * Math.Pow(10, digits)) / Math.Pow(10, digits);

            return Math.Ceiling(value * Math.Pow(10, digits)) / Math.Pow(10, digits);
        }

        private static object RoundUp(List<Expression> p)
        {
            var value = (Double)p[0];
            var digits = (Int32)(Double)p[1];

            if (value >= 0)
                return Math.Ceiling(value * Math.Pow(10, digits)) / Math.Pow(10, digits);

            return Math.Floor(value * Math.Pow(10, digits)) / Math.Pow(10, digits);
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

        private static object SumIf(List<Expression> p)
        {
            // get parameters
            if (!CalcEngineHelpers.TryExtractRange(p[0], out var range, out var calculationErrorType))
            {
                return calculationErrorType;
            }

            // range of values to match the criteria against
            // limit to first column only
            var rangeColumn = new CellRangeReference(range!.Column(1).AsRange()) as IEnumerable;

            // range of values to sum up
            var sumRange = p.Count < 3 ?
                p[0] as XObjectExpression :
                p[2] as XObjectExpression;

            // the criteria to evaluate
            var criteria = p[1].Evaluate();

            var rangeValues = rangeColumn.Cast<object>().ToList();
            using var sumRangeEnumerator = sumRange.Cast<object>().GetEnumerator();

            // compute total
            var ce = new XLCalcEngine(CultureInfo.CurrentCulture);
            var tally = new Tally();
            for (var i = 0; i < rangeValues.Count; i++)
            {
                // TODO: Replace this mess completely
                var targetValue = rangeValues[i];
                if (CalcEngineHelpers.ValueSatisfiesCriteria(targetValue, criteria, ce))
                {
                    if (!sumRangeEnumerator.MoveNext())
                        break;
                    var value = sumRangeEnumerator.Current!;
                    tally.AddValue(value);
                }
                else
                {
                    try
                    {
                        if (!sumRangeEnumerator.MoveNext())
                            break;
                    }
                    catch (GettingDataException)
                    {
                        // The referenced cell uses a dirty formula, but we are not using the value, so eat the exception.
                    }
                }
            }

            // done
            return tally.Sum();
        }

        private static object SumIfs(List<Expression> p)
        {
            // get parameters
            var sumRange = (IEnumerable)p[0];
            var sumRangeDimensions = CalcEngineHelpers.GetRangeDimensions(p[0] as XObjectExpression);

            var sumRangeValues = new List<object>();
            foreach (var value in sumRange)
            {
                sumRangeValues.Add(value);
            }

            var ce = new XLCalcEngine(CultureInfo.CurrentCulture);
            var tally = new Tally();

            int numberOfCriteria = p.Count / 2; // int division returns floor() automatically, that's what we want.

            for (int criteriaPair = 0; criteriaPair < numberOfCriteria; criteriaPair++)
            {
                var criterionDimensions = CalcEngineHelpers.GetRangeDimensions(p[criteriaPair * 2 + 1] as XObjectExpression);
                if (criterionDimensions != sumRangeDimensions)
                {
                    return XLError.IncompatibleValue;
                }
            }

            // prepare criteria-parameters:
            var criteriaRanges = new Tuple<object, IList<object>>[numberOfCriteria];
            for (int criteriaPair = 0; criteriaPair < numberOfCriteria; criteriaPair++)
            {
                if (p[criteriaPair * 2 + 1] is IEnumerable criteriaRange)
                {
                    var criterion = p[criteriaPair * 2 + 2].Evaluate();
                    var criteriaRangeValues = criteriaRange.Cast<Object>().ToList();

                    criteriaRanges[criteriaPair] = new Tuple<object, IList<object>>(
                        criterion,
                        criteriaRangeValues);
                }
                else
                {
                    return XLError.CellReference;
                }
            }

            for (var i = 0; i < sumRangeValues.Count; i++)
            {
                bool shouldUseValue = true;

                foreach (var criteriaPair in criteriaRanges)
                {
                    if (!CalcEngineHelpers.ValueSatisfiesCriteria(
                        i < criteriaPair.Item2.Count ? criteriaPair.Item2[i] : string.Empty,
                        criteriaPair.Item1,
                        ce))
                    {
                        shouldUseValue = false;
                        break; // we're done with the inner loop as we can't ever get true again.
                    }
                }

                if (shouldUseValue)
                    tally.AddValue(sumRangeValues[i]);
            }

            // done
            return tally.Sum();
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
