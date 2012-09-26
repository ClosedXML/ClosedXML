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
    }
}
