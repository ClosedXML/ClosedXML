using System;
using System.Text;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    public static class XLMath
    {
        public static double DegreesToRadians(double degrees)
        {
            return Math.PI / 180.0 * degrees;
        }

        public static double RadiansToDegrees(double radians)
        {
            return 180.0 / Math.PI * radians;
        }

        public static double GradsToRadians(double grads)
        {
            return grads / 200.0 * Math.PI;
        }

        public static double RadiansToGrads(double radians)
        {
            return radians / Math.PI * 200.0;
        }

        public static double DegreesToGrads(double degrees)
        {
            return degrees / 9.0 * 10.0;
        }

        public static double GradsToDegrees(double grads)
        {
            return grads / 10.0 * 9.0;
        }

        public static double ASinh(double x)
        {
            return Math.Log(x + Math.Sqrt(x * x + 1.0));
        }

        public static double ACosh(double x)
        {
            return Math.Log(x + Math.Sqrt((x * x) - 1.0));
        }

        public static double ATanh(double x)
        {
            return Math.Log((1.0 + x) / (1.0 - x)) / 2.0;
        }

        public static double ACoth(double x)
        {
            //return (Math.Log((x + 1.0) / (x - 1.0)) / 2.0);
            return ATanh(1.0 / x);
        }

        public static double ASech(double x)
        {
            return ACosh(1.0 / x);
        }

        public static double ACsch(double x)
        {
            return ASinh(1.0 / x);
        }

        public static double Sech(double x)
        {
            return 1.0 / Math.Cosh(x);
        }

        public static double Csch(double x)
        {
            return 1.0 / Math.Sinh(x);
        }

        public static double Coth(double x)
        {
            return Math.Cosh(x) / Math.Sinh(x);
        }

        public static double Combin(int n, int k)
        {
            if (k == 0)
            {
                return 1;
            }

            return n * Combin(n - 1, k - 1) / k;
        }

        public static bool IsEven(int value)
        {
            return Math.Abs(value % 2) == 0;
        }
        public static bool IsOdd(int value)
        {
            return Math.Abs(value % 2) != 0;
        }

        public static string ToRoman(int number)
        {
            if ((number < 0) || (number > 3999))
            {
                throw new ArgumentOutOfRangeException("insert value betwheen 1 and 3999");
            }

            if (number < 1)
            {
                return string.Empty;
            }

            if (number >= 1000)
            {
                return "M" + ToRoman(number - 1000);
            }

            if (number >= 900)
            {
                return "CM" + ToRoman(number - 900);
            }

            if (number >= 500)
            {
                return "D" + ToRoman(number - 500);
            }

            if (number >= 400)
            {
                return "CD" + ToRoman(number - 400);
            }

            if (number >= 100)
            {
                return "C" + ToRoman(number - 100);
            }

            if (number >= 90)
            {
                return "XC" + ToRoman(number - 90);
            }

            if (number >= 50)
            {
                return "L" + ToRoman(number - 50);
            }

            if (number >= 40)
            {
                return "XL" + ToRoman(number - 40);
            }

            if (number >= 10)
            {
                return "X" + ToRoman(number - 10);
            }

            if (number >= 9)
            {
                return "IX" + ToRoman(number - 9);
            }

            if (number >= 5)
            {
                return "V" + ToRoman(number - 5);
            }

            if (number >= 4)
            {
                return "IV" + ToRoman(number - 4);
            }

            if (number >= 1)
            {
                return "I" + ToRoman(number - 1);
            }

            throw new ArgumentOutOfRangeException("something bad happened");
        }

        public static int RomanToArabic(string text)
        {
            if (text.Length == 0)
            {
                return 0;
            }

            if (text.StartsWith("M", StringComparison.InvariantCultureIgnoreCase))
            {
                return 1000 + RomanToArabic(text.Substring(1));
            }

            if (text.StartsWith("CM", StringComparison.InvariantCultureIgnoreCase))
            {
                return 900 + RomanToArabic(text.Substring(2));
            }

            if (text.StartsWith("D", StringComparison.InvariantCultureIgnoreCase))
            {
                return 500 + RomanToArabic(text.Substring(1));
            }

            if (text.StartsWith("CD", StringComparison.InvariantCultureIgnoreCase))
            {
                return 400 + RomanToArabic(text.Substring(2));
            }

            if (text.StartsWith("C", StringComparison.InvariantCultureIgnoreCase))
            {
                return 100 + RomanToArabic(text.Substring(1));
            }

            if (text.StartsWith("XC", StringComparison.InvariantCultureIgnoreCase))
            {
                return 90 + RomanToArabic(text.Substring(2));
            }

            if (text.StartsWith("L", StringComparison.InvariantCultureIgnoreCase))
            {
                return 50 + RomanToArabic(text.Substring(1));
            }

            if (text.StartsWith("XL", StringComparison.InvariantCultureIgnoreCase))
            {
                return 40 + RomanToArabic(text.Substring(2));
            }

            if (text.StartsWith("X", StringComparison.InvariantCultureIgnoreCase))
            {
                return 10 + RomanToArabic(text.Substring(1));
            }

            if (text.StartsWith("IX", StringComparison.InvariantCultureIgnoreCase))
            {
                return 9 + RomanToArabic(text.Substring(2));
            }

            if (text.StartsWith("V", StringComparison.InvariantCultureIgnoreCase))
            {
                return 5 + RomanToArabic(text.Substring(1));
            }

            if (text.StartsWith("IV", StringComparison.InvariantCultureIgnoreCase))
            {
                return 4 + RomanToArabic(text.Substring(2));
            }

            if (text.StartsWith("I", StringComparison.InvariantCultureIgnoreCase))
            {
                return 1 + RomanToArabic(text.Substring(1));
            }

            throw new ArgumentOutOfRangeException("text is not a valid roman number");
        }

        public static string ChangeBase(long number, int radix)
        {
            if (number < 0)
            {
                throw new ArgumentOutOfRangeException("number must be greater or equal to 0");
            }

            if (radix < 2)
            {
                throw new ArgumentOutOfRangeException("radix must be greater or equal to 2");
            }

            if (radix > 36)
            {
                throw new ArgumentOutOfRangeException("radix must be smaller than or equal to 36");
            }

            var sb = new StringBuilder();
            var remaining = number;

            if (remaining == 0)
            {
                sb.Insert(0, '0');
            }

            while (remaining > 0)
            {
                var nextDigitDecimal = remaining % radix;
                remaining = remaining / radix;

                if (nextDigitDecimal < 10)
                {
                    sb.Insert(0, nextDigitDecimal);
                }
                else
                {
                    sb.Insert(0, (char)(nextDigitDecimal + 55));
                }
            }

            return sb.ToString();
        }
    }
}
