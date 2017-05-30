using System;
using System.Collections.Generic;
using System.Linq;
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

        public static double ASinh(double x)
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
            return (ASinh(1.0 / x));
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

        public static double Combin(Int32 n, Int32 k)
        {
            if (k == 0) return 1;
            return n * Combin(n - 1, k - 1) / k;
        }

        public static Boolean IsEven(Int32 value)
        {
            return Math.Abs(value % 2) < XLHelper.Epsilon;
        }
        public static Boolean IsOdd(Int32 value)
        {
            return Math.Abs(value % 2) > XLHelper.Epsilon;
        }

        public static string ToRoman(int number)
        {
            if ((number < 0) || (number > 3999)) throw new ArgumentOutOfRangeException("insert value betwheen 1 and 3999");
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + ToRoman(number - 1000);
            if (number >= 900) return "CM" + ToRoman(number - 900); 
            if (number >= 500) return "D" + ToRoman(number - 500);
            if (number >= 400) return "CD" + ToRoman(number - 400);
            if (number >= 100) return "C" + ToRoman(number - 100);
            if (number >= 90) return "XC" + ToRoman(number - 90);
            if (number >= 50) return "L" + ToRoman(number - 50);
            if (number >= 40) return "XL" + ToRoman(number - 40);
            if (number >= 10) return "X" + ToRoman(number - 10);
            if (number >= 9) return "IX" + ToRoman(number - 9);
            if (number >= 5) return "V" + ToRoman(number - 5);
            if (number >= 4) return "IV" + ToRoman(number - 4);
            if (number >= 1) return "I" + ToRoman(number - 1);
            throw new ArgumentOutOfRangeException("something bad happened");
        }

    }
}
