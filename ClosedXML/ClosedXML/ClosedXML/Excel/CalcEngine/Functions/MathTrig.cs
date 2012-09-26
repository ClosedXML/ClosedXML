using System;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel.CalcEngine.Functions;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class MathTrig
    {
        private static readonly Random _rnd = new Random();

        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("ABS", 1, Abs);
            ce.RegisterFunction("ACOS", 1, Acos);
            ce.RegisterFunction("ACOSH", 1, Acosh);
            ce.RegisterFunction("ASIN", 1, Asin);
            ce.RegisterFunction("ASINH", 1, Asinh);
            ce.RegisterFunction("ATAN", 1, Atan);
            ce.RegisterFunction("ATAN2", 2, Atan2);
            ce.RegisterFunction("ATANH", 1, Atanh);
            ce.RegisterFunction("CEILING", 1, Ceiling);
            ce.RegisterFunction("COMBIN", 2, Combin);
            ce.RegisterFunction("COS", 1, Cos);
            ce.RegisterFunction("COSH", 1, Cosh);
            ce.RegisterFunction("DEGREES", 1, Degrees);
            //ce.RegisterFunction("EVEN", Even, 1);
            ce.RegisterFunction("EXP", 1, Exp);
            //ce.RegisterFunction("FACT", Fact, 1);
            //ce.RegisterFunction("FACTDOUBLE", FactDouble, 1);
            ce.RegisterFunction("FLOOR", 1, Floor);
            //ce.RegisterFunction("GCD", Gcd, 1);
            ce.RegisterFunction("INT", 1, Int);
            //ce.RegisterFunction("LCM", Lcm, 1);
            ce.RegisterFunction("LN", 1, Ln);
            ce.RegisterFunction("LOG", 1, 2, Log);
            ce.RegisterFunction("LOG10", 1, Log10);
            //ce.RegisterFunction("MDETERM", MDeterm, 1);
            //ce.RegisterFunction("MINVERSE", MInverse, 1);
            //ce.RegisterFunction("MMULT", MMult, 1);
            //ce.RegisterFunction("MOD", Mod, 2);
            //ce.RegisterFunction("MROUND", MRound, 1);
            //ce.RegisterFunction("MULTINOMIAL", Multinomial, 1);
            //ce.RegisterFunction("ODD", Odd, 1);
            ce.RegisterFunction("PI", 0, Pi);
            ce.RegisterFunction("POWER", 2, Power);
            //ce.RegisterFunction("PRODUCT", Product, 1);
            //ce.RegisterFunction("QUOTIENT", Quotient, 1);
            //ce.RegisterFunction("RADIANS", Radians, 1);
            ce.RegisterFunction("RAND", 0, Rand);
            ce.RegisterFunction("RANDBETWEEN", 2, RandBetween);
            //ce.RegisterFunction("ROMAN", Roman, 1);
            //ce.RegisterFunction("ROUND", Round, 1);
            //ce.RegisterFunction("ROUNDDOWN", RoundDown, 1);
            //ce.RegisterFunction("ROUNDUP", RoundUp, 1);
            //ce.RegisterFunction("SERIESSUM", SeriesSum, 1);
            ce.RegisterFunction("SIGN", 1, Sign);
            ce.RegisterFunction("SIN", 1, Sin);
            ce.RegisterFunction("SINH", 1, Sinh);
            ce.RegisterFunction("SQRT", 1, Sqrt);
            //ce.RegisterFunction("SQRTPI", SqrtPi, 1);
            //ce.RegisterFunction("SUBTOTAL", Subtotal, 1);
            ce.RegisterFunction("SUM", 1, int.MaxValue, Sum);
            ce.RegisterFunction("SUMIF", 2, 3, SumIf);
            //ce.RegisterFunction("SUMPRODUCT", SumProduct, 1);
            //ce.RegisterFunction("SUMSQ", SumSq, 1);
            //ce.RegisterFunction("SUMX2MY2", SumX2MY2, 1);
            //ce.RegisterFunction("SUMX2PY2", SumX2PY2, 1);
            //ce.RegisterFunction("SUMXMY2", SumXMY2, 1);
            ce.RegisterFunction("TAN", 1, Tan);
            ce.RegisterFunction("TANH", 1, Tanh);
            ce.RegisterFunction("TRUNC", 1, Trunc);
        }

        private static object Abs(List<Expression> p)
        {
            return Math.Abs(p[0]);
        }

        private static object Acos(List<Expression> p)
        {
            return Math.Acos(p[0]);
        }

        private static object Asin(List<Expression> p)
        {
            return Math.Asin(p[0]);
        }

        private static object Atan(List<Expression> p)
        {
            return Math.Atan(p[0]);
        }

        private static object Atan2(List<Expression> p)
        {
            return Math.Atan2(p[0], p[1]);
        }

        private static object Ceiling(List<Expression> p)
        {
            return Math.Ceiling(p[0]);
        }

        private static object Cos(List<Expression> p)
        {
            return Math.Cos(p[0]);
        }

        private static object Cosh(List<Expression> p)
        {
            return Math.Cosh(p[0]);
        }

        private static object Exp(List<Expression> p)
        {
            return Math.Exp(p[0]);
        }

        private static object Floor(List<Expression> p)
        {
            return Math.Floor(p[0]);
        }

        private static object Int(List<Expression> p)
        {
            return (int) ((double) p[0]);
        }

        private static object Ln(List<Expression> p)
        {
            return Math.Log(p[0]);
        }

        private static object Log(List<Expression> p)
        {
            var lbase = p.Count > 1 ? (double) p[1] : 10;
            return Math.Log(p[0], lbase);
        }

        private static object Log10(List<Expression> p)
        {
            return Math.Log10(p[0]);
        }

        private static object Pi(List<Expression> p)
        {
            return Math.PI;
        }

        private static object Power(List<Expression> p)
        {
            return Math.Pow(p[0], p[1]);
        }

        private static object Rand(List<Expression> p)
        {
            return _rnd.NextDouble();
        }

        private static object RandBetween(List<Expression> p)
        {
            return _rnd.Next((int) (double) p[0], (int) (double) p[1]);
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
            var range = p[0] as IEnumerable;
            var sumRange = p.Count < 3 ? range : p[2] as IEnumerable;
            var criteria = p[1].Evaluate();

            // build list of values in range and sumRange
            var rangeValues = new List<object>();
            foreach (var value in range)
            {
                rangeValues.Add(value);
            }
            var sumRangeValues = new List<object>();
            foreach (var value in sumRange)
            {
                sumRangeValues.Add(value);
            }

            // compute total
            var ce = new CalcEngine();
            var tally = new Tally();
            for (var i = 0; i < Math.Min(rangeValues.Count, sumRangeValues.Count); i++)
            {
                if (ValueSatisfiesCriteria(rangeValues[i], criteria, ce))
                {
                    tally.AddValue(sumRangeValues[i]);
                }
            }

            // done
            return tally.Sum();
        }

        private static bool ValueSatisfiesCriteria(object value, object criteria, CalcEngine ce)
        {
            // safety...
            if (value == null)
            {
                return false;
            }

            // if criteria is a number, straight comparison
            if (criteria is double)
            {
                return (double) value == (double) criteria;
            }

            // convert criteria to string
            var cs = criteria as string;
            if (!string.IsNullOrEmpty(cs))
            {
                // if criteria is an expression (e.g. ">20"), use calc engine
                if (cs[0] == '=' || cs[0] == '<' || cs[0] == '>')
                {
                    // build expression
                    var expression = string.Format("{0}{1}", value, cs);

                    // add quotes if necessary
                    var pattern = @"(\w+)(\W+)(\w+)";
                    var m = Regex.Match(expression, pattern);
                    if (m.Groups.Count == 4)
                    {
                        double d;
                        if (!double.TryParse(m.Groups[1].Value, out d) ||
                            !double.TryParse(m.Groups[3].Value, out d))
                        {
                            expression = string.Format("\"{0}\"{1}\"{2}\"",
                                                       m.Groups[1].Value,
                                                       m.Groups[2].Value,
                                                       m.Groups[3].Value);
                        }
                    }

                    // evaluate
                    return (bool) ce.Evaluate(expression);
                }

                // if criteria is a regular expression, use regex
                if (cs.IndexOf('*') > -1)
                {
                    var pattern = cs.Replace(@"\", @"\\");
                    pattern = pattern.Replace(".", @"\");
                    pattern = pattern.Replace("*", ".*");
                    return Regex.IsMatch(value.ToString(), pattern, RegexOptions.IgnoreCase);
                }

                // straight string comparison 
                return string.Equals(value.ToString(), cs, StringComparison.OrdinalIgnoreCase);
            }

            // should never get here?
            Debug.Assert(false, "failed to evaluate criteria in SumIf");
            return false;
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
            return (double) (int) ((double) p[0]);
        }

        public static double DegreesToRadians(double degrees)
        {
            return (Math.PI/180.0)*degrees;
        }

        public static double RadiansToDegrees(double radians)
        {
            return (180.0/Math.PI)*radians;
        }

        public static double GradsToRadians(double grads)
        {
            return (grads/200.0)*Math.PI;
        }

        public static double RadiansToGrads(double radians)
        {
            return (radians/Math.PI)*200.0;
        }

        public static double DegreesToGrads(double degrees)
        {
            return (degrees/9.0)*10.0;
        }

        public static double GradsToDegrees(double grads)
        {
            return (grads/10.0)*9.0;
        }

        public static double ASinh(double x)
        {
            return (Math.Log(x + Math.Sqrt(x*x + 1.0)));
        }

        private static object Acosh(List<Expression> p)
        {
            return XLMath.ACosh(p[0]);
        }

        private static object Asinh(List<Expression> p)
        {
            return XLMath.ASinh(p[0]);
        }

        private static object Atanh(List<Expression> p)
        {
            return XLMath.ATanh(p[0]);
        }

        private static object Combin(List<Expression> p)
        {
            Int32 n = (int) p[0];
            Int32 k = (int) p[1];
            return XLMath.Combin(n, k);
        }

        private static object Degrees(List<Expression> p)
        {
            return Math.PI * p[0] / 180.0;
        }
    }
}