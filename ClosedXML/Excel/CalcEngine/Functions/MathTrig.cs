// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel.CalcEngine.Exceptions;
using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class MathTrig
    {
        private static readonly Random _rnd = new Random();

        #region Register

        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("ABS", 1, Abs);
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
            ce.RegisterFunction("COMBIN", 2, Combin);
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
            ce.RegisterFunction("FACT", 1, Fact);
            ce.RegisterFunction("FACTDOUBLE", 1, FactDouble);
            ce.RegisterFunction("FLOOR", 2, Floor);
            ce.RegisterFunction("FLOOR.MATH", 1, 3, FloorMath);
            ce.RegisterFunction("GCD", 1, 255, Gcd);
            ce.RegisterFunction("INT", 1, Int);
            ce.RegisterFunction("LCM", 1, 255, Lcm);
            ce.RegisterFunction("LN", 1, Ln);
            ce.RegisterFunction("LOG", 1, 2, Log);
            ce.RegisterFunction("LOG10", 1, Log10);
            ce.RegisterFunction("MDETERM", 1, MDeterm);
            ce.RegisterFunction("MINVERSE", 1, MInverse);
            ce.RegisterFunction("MMULT", 2, MMult);
            ce.RegisterFunction("MOD", 2, Mod);
            ce.RegisterFunction("MROUND", 2, MRound);
            ce.RegisterFunction("MULTINOMIAL", 1, 255, Multinomial);
            ce.RegisterFunction("ODD", 1, Odd);
            ce.RegisterFunction("PI", 0, Pi);
            ce.RegisterFunction("POWER", 2, Power);
            ce.RegisterFunction("PRODUCT", 1, 255, Product);
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
            ce.RegisterFunction("SERIESSUM", 4, SeriesSum);
            ce.RegisterFunction("SIGN", 1, Sign);
            ce.RegisterFunction("SIN", 1, Sin);
            ce.RegisterFunction("SINH", 1, Sinh);
            ce.RegisterFunction("SQRT", 1, Sqrt);
            ce.RegisterFunction("SQRTPI", 1, SqrtPi);
            ce.RegisterFunction("SUBTOTAL", 2, 255, Subtotal);
            ce.RegisterFunction("SUM", 1, int.MaxValue, Sum);
            ce.RegisterFunction("SUMIF", 2, 3, SumIf);
            ce.RegisterFunction("SUMIFS", 3, 255, SumIfs);
            ce.RegisterFunction("SUMPRODUCT", 1, 30, SumProduct);
            ce.RegisterFunction("SUMSQ", 1, 255, SumSq);
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

        private static object Abs(List<Expression> p)
        {
            return Math.Abs(p[0]);
        }

        private static object Acos(List<Expression> p)
        {
            double input = p[0];
            if (Math.Abs(input) > 1)
                throw new NumberException();

            return Math.Acos(p[0]);
        }

        private static object Acosh(List<Expression> p)
        {
            double number = p[0];
            if (number < 1)
                throw new NumberException();

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
                throw new NumberException();

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
                    throw new NumberException();
                else if (input[0] == '-')
                    return -XLMath.RomanToArabic(input.Substring(1));
                else
                    return XLMath.RomanToArabic(input);
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new CellValueException();
            }
        }

        private static object Asin(List<Expression> p)
        {
            double input = p[0];
            if (Math.Abs(input) > 1)
                throw new NumberException();

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
                throw new DivisionByZeroException();

            return Math.Atan2(y, x);
        }

        private static object Atanh(List<Expression> p)
        {
            double input = p[0];
            if (Math.Abs(input) >= 1)
                throw new NumberException();

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
                throw new CellValueException();

            var rawRadix = p[1].Evaluate();
            if (rawRadix is long || rawRadix is int || rawRadix is byte || rawRadix is double || rawRadix is float)
                radix = Convert.ToInt32(rawRadix);
            else
                throw new CellValueException();

            if (p.Count > 2)
            {
                var rawMinLength = p[2].Evaluate();
                if (rawMinLength is long || rawMinLength is int || rawMinLength is byte || rawMinLength is double || rawMinLength is float)
                    minLength = Convert.ToInt32(rawMinLength);
                else
                    throw new CellValueException();
            }

            if (number < 0 || radix < 2 || radix > 36)
                throw new NumberException();

            return XLMath.ChangeBase(number, radix).PadLeft(minLength, '0');
        }

        private static object Ceiling(List<Expression> p)
        {
            double number = p[0];
            double significance = p[1];

            if (significance == 0)
                return 0d;
            else if (significance < 0 && number > 0)
                throw new NumberException();
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

        private static object Combin(List<Expression> p)
        {
            Int32 n;
            Int32 k;

            var rawN = p[0].Evaluate();
            var rawK = p[1].Evaluate();
            if (rawN is long || rawN is int || rawN is byte || rawN is double || rawN is float)
                n = (int)Math.Floor((double)rawN);
            else
                throw new NumberException();

            if (rawK is long || rawK is int || rawK is byte || rawK is double || rawK is float)
                k = (int)Math.Floor((double)rawK);
            else
                throw new NumberException();

            n = (int)p[0];
            k = (int)p[1];

            if (n < 0 || n < k || k < 0)
                throw new NumberException();

            return XLMath.Combin(n, k);
        }

        private static object CombinA(List<Expression> p)
        {
            Int32 number = (int)p[0]; // casting truncates towards 0 as specified
            Int32 chosen = (int)p[1];

            if (number < 0 || number < chosen)
                throw new NumberException();
            if (chosen < 0)
                throw new NumberException();

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
                throw new DivisionByZeroException();

            return 1 / tan;
        }

        private static object Coth(List<Expression> p)
        {
            double input = p[0];
            if (input == 0)
                throw new DivisionByZeroException();

            return 1 / Math.Tanh(input);
        }

        private static object Csc(List<Expression> p)
        {
            double input = p[0];
            if (input == 0)
                throw new DivisionByZeroException();

            return 1 / Math.Sin(input);
        }

        private static object Csch(List<Expression> p)
        {
            if (Math.Abs((double)p[0].Evaluate()) < Double.Epsilon)
                throw new DivisionByZeroException();

            return 1 / Math.Sinh(p[0]);
        }

        private static object Decimal(List<Expression> p)
        {
            string source = p[0];
            double radix = p[1];

            if (radix < 2 || radix > 36)
                throw new NumberException();

            var asciiValues = Encoding.ASCII.GetBytes(source.ToUpperInvariant());

            double result = 0;
            int i = 0;

            foreach (byte digit in asciiValues)
            {
                if (digit > 90)
                {
                    throw new NumberException();
                }

                int digitNumber = digit >= 48 && digit < 58
                    ? digit - 48
                    : digit - 55;

                if (digitNumber > radix - 1)
                    throw new NumberException();

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

        private static object Fact(List<Expression> p)
        {
            var input = p[0].Evaluate();

            if (!(input is long || input is int || input is byte || input is double || input is float))
                throw new CellValueException();

            var num = Math.Floor((double)input);
            double fact = 1.0;

            if (num < 0)
                throw new NumberException();

            if (num > 1)
            {
                for (int i = 2; i <= num; i++)
                {
                    fact *= i;
                }
            }

            return fact;
        }

        private static object FactDouble(List<Expression> p)
        {
            var input = p[0].Evaluate();

            if (!(input is long || input is int || input is byte || input is double || input is float))
                throw new CellValueException();

            var num = Math.Floor(p[0]);
            double fact = 1.0;

            if (num < -1)
                throw new NumberException();

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
                throw new DivisionByZeroException();
            else if (significance < 0 && number > 0)
                throw new NumberException();
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
            catch (FormatException e)
            {
                throw new CellValueException("Cells are empty or contain text.", e);
            }

            if (A.GetLength(1) != B.GetLength(0))
                throw new CellValueException("The number of columns in array1 is different from the number of rows in array2.");

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
                throw new NumberException("The Number and Multiple arguments must have the same sign.");

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

        private static object Product(List<Expression> p)
        {
            if (p.Count == 0) return 0;
            Double total = 1;
            p.ForEach(v => total *= v);
            return total;
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
                throw new CellValueException();
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

        private static object Subtotal(List<Expression> p)
        {
            // Skip cells that already evaluate a SUBTOTAL
            bool hasSubtotalInFormula(Expression e)
            {
                if (e is FunctionExpression fe && (fe.FunctionDefinition.Function.Method.Name == nameof(Subtotal) || fe.Parameters.Any(fp => hasSubtotalInFormula(fp))))
                    return true;

                if (e is BinaryExpression be)
                    return hasSubtotalInFormula(be.LeftExpression) || hasSubtotalInFormula(be.RightExpression);

                if (e is UnaryExpression ue)
                    return hasSubtotalInFormula(ue.Expression);

                return false;
            }

            IEnumerable<Expression> extractExpressionsWithoutSubtotal(CellRangeReference crr)
            {
                var ce = crr.CalcEngine as XLCalcEngine;

                return crr.Range
                    .CellsUsed()
                    .Where(c =>
                    {
                        if (c.HasFormula)
                        {
                            var expression = ce.ExpressionCache[c.FormulaA1];
                            return !hasSubtotalInFormula(expression);
                        }
                        else
                        {
                            return true;
                        }
                    })
                    .Select(c => new XObjectExpression(new CellRangeReference(c.AsRange(), (XLCalcEngine)crr.CalcEngine)) as Expression);
            }

            var expressions = p.Skip(1)
                .SelectMany(e =>
                    e is XObjectExpression xoe && xoe.Value is CellRangeReference crr
                        ? extractExpressionsWithoutSubtotal(crr)
                        : new[] { e })
                .ToArray();

            var fId = (int)(Double)p[0];
            var tally = new Tally(expressions);

            return fId switch
            {
                1 => tally.Average(),
                2 => tally.Count(true),
                3 => tally.Count(false),
                4 => tally.Max(),
                5 => tally.Min(),
                6 => tally.Product(),
                7 => tally.Std(),
                8 => tally.StdP(),
                9 => tally.Sum(),
                10 => tally.Var(),
                11 => tally.VarP(),
                _ => throw new ArgumentException("Function not supported."),
            };
        }

        private static object Sum(List<Expression> p)
        {
            var tally = new Tally();
            foreach (var e in p)
            {
                tally.Add(e);
            }
            return tally.Sum();
        }

        private static object SumIf(List<Expression> p)
        {
            // get parameters
            var range = p[0] as IEnumerable;                            // range of values to match the criteria against
            var sumRange = p.Count < 3 ?
                p[0] as XObjectExpression :
                p[2] as XObjectExpression;   // range of values to sum up
            var criteria = p[1].Evaluate();                             // the criteria to evaluate

            var rangeValues = range.Cast<object>().ToList();
            var sumRangeValues = sumRange.Cast<object>().ToList();

            // compute total
            var ce = new CalcEngine();
            var tally = new Tally();
            for (var i = 0; i < Math.Max(rangeValues.Count, sumRangeValues.Count); i++)
            {
                var targetValue = i < rangeValues.Count ? rangeValues[i] : string.Empty;
                if (CalcEngineHelpers.ValueSatisfiesCriteria(targetValue, criteria, ce))
                {
                    var value = i < sumRangeValues.Count ? sumRangeValues[i] : 0d;
                    tally.AddValue(value);
                }
            }

            // done
            return tally.Sum();
        }

        private static object SumIfs(List<Expression> p)
        {
            // get parameters
            var sumRange = p[0] as IEnumerable;

            var sumRangeValues = new List<object>();
            foreach (var value in sumRange)
            {
                sumRangeValues.Add(value);
            }

            var ce = new CalcEngine();
            var tally = new Tally();

            int numberOfCriteria = p.Count / 2; // int division returns floor() automatically, that's what we want.

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
                    throw new CellReferenceException($"Expected parameter {criteriaPair * 2 + 2} to be a range");
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

        private static object SumProduct(List<Expression> p)
        {
            // all parameters should be IEnumerable
            if (p.Any(param => !(param is IEnumerable)))
                throw new NoValueAvailableException();

            var counts = p.Cast<IEnumerable>().Select(param =>
            {
                int i = 0;
                foreach (var item in param)
                    i++;
                return i;
            })
            .Distinct();

            // All parameters should have the same length
            if (counts.Count() > 1)
                throw new NoValueAvailableException();

            var values = p
                .Cast<IEnumerable>()
                .Select(range =>
                {
                    var results = new List<double>();
                    foreach (var c in range)
                    {
                        if (c.IsNumber())
                            results.Add(c.CastTo<double>());
                        else
                            results.Add(0.0);
                    }
                    return results;
                })
                .ToArray();

            return Enumerable.Range(0, counts.Single())
                .Aggregate(0d, (t, i) =>
                    t + values.Aggregate(1d,
                        (product, list) => product * list[i]
                    )
                );
        }

        private static object SumSq(List<Expression> p)
        {
            var t = new Tally(p);
            return t.NumericValues().Sum(v => Math.Pow(v, 2));
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
    }
}
