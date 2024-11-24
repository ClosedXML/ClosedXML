using System;
using System.Text;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    public static class XLMath
    {
        public static double DegreesToRadians(double degrees)
        {
            return (Math.PI / 180.0) * degrees;
        }

        public static double RadiansToDegrees(double radians)
        {
            return (180.0 / Math.PI) * radians;
        }

        public static double GradsToRadians(double grads)
        {
            return (grads / 200.0) * Math.PI;
        }

        public static double RadiansToGrads(double radians)
        {
            return (radians / Math.PI) * 200.0;
        }

        public static double DegreesToGrads(double degrees)
        {
            return (degrees / 9.0) * 10.0;
        }

        public static double GradsToDegrees(double grads)
        {
            return (grads / 10.0) * 9.0;
        }

        public static double Asinh(double x)
        {
            return (Math.Log(x + Math.Sqrt(x * x + 1.0)));
        }

        public static double ACosh(double x)
        {
            return (Math.Log(x + Math.Sqrt((x * x) - 1.0)));
        }

        public static double ATanh(double x)
        {
            return (Math.Log((1.0 + x) / (1.0 - x)) / 2.0);
        }

        public static double ACoth(double x)
        {
            //return (Math.Log((x + 1.0) / (x - 1.0)) / 2.0);
            return (ATanh(1.0 / x));
        }

        public static double ASech(double x)
        {
            return (ACosh(1.0 / x));
        }

        public static double ACsch(double x)
        {
            return (Asinh(1.0 / x));
        }

        public static double Sech(double x)
        {
            return (1.0 / Math.Cosh(x));
        }

        public static double Csch(double x)
        {
            return (1.0 / Math.Sinh(x));
        }

        public static double Coth(double x)
        {
            return (Math.Cosh(x) / Math.Sinh(x));
        }

        internal static OneOf<double, XLError> CombinChecked(double number, double numberChosen)
        {
            if (number < 0 || numberChosen < 0)
                return XLError.NumberInvalid;

            var n = Math.Floor(number);
            var k = Math.Floor(numberChosen);

            // Parameter doesn't fit into int. That's how many multiplications Excel allows.
            if (n >= int.MaxValue || k >= int.MaxValue)
                return XLError.NumberInvalid;

            if (n < k)
                return XLError.NumberInvalid;

            var combinations = Combin(n, k);
            if (double.IsInfinity(combinations) || double.IsNaN(combinations))
                return XLError.NumberInvalid;

            return combinations;
        }

        internal static double Combin(double n, double k)
        {
            if (k == 0) return 1;

            // Don't use recursion, malicious input could exhaust stack.
            // Don't calculate directly from factorials, could overflow.
            double result = 1;
            for (var i = 1; i <= k; i++, n--)
            {
                result *= n;
                result /= i;
            }

            return result;
        }

        internal static double Factorial(double n)
        {
            n = Math.Truncate(n);
            var factorial = 1d;
            while (n > 1)
            {
                factorial *= n--;

                // n can be very large, stop when we reach infinity.
                if (double.IsInfinity(factorial))
                    return factorial;
            }

            return factorial;
        }

        public static Boolean IsEven(Int32 value)
        {
            return Math.Abs(value % 2) == 0;
        }

        public static Boolean IsEven(double value)
        {
            // Check the number doesn't have any fractions and that it is even.
            // Due to rounding after division, only checking for % 2 could fail
            // for numbers really close to whole number.
            var hasNoFraction = value % 1 == 0;
            var isEven = value % 2 == 0;
            return hasNoFraction && isEven;
        }

        public static Boolean IsOdd(Int32 value)
        {
            return Math.Abs(value % 2) != 0;
        }

        public static Boolean IsOdd(double value)
        {
            var hasNoFraction = value % 1 == 0;
            var isOdd = value % 2 != 0;
            return hasNoFraction && isOdd;
        }
    }
}
