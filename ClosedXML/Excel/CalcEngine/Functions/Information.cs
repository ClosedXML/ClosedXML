using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class Information
    {
        public static void Register(CalcEngine ce)
        {
            //TODO: Add documentation
            ce.RegisterFunction("ERROR.TYPE", 1, ErrorType);
            ce.RegisterFunction("ISBLANK", 1, 1, IsBlank);
            ce.RegisterFunction("ISERR", 1, 1, IsErr);
            ce.RegisterFunction("ISERROR", 1, 1, IsError);
            ce.RegisterFunction("ISEVEN", 1, IsEven);
            ce.RegisterFunction("ISLOGICAL", 1, int.MaxValue, IsLogical);
            ce.RegisterFunction("ISNA", 1, 1, IsNa);
            ce.RegisterFunction("ISNONTEXT", 1, int.MaxValue, IsNonText);
            ce.RegisterFunction("ISNUMBER", 1, int.MaxValue, IsNumber);
            ce.RegisterFunction("ISODD", 1, IsOdd);
            ce.RegisterFunction("ISREF", 1, int.MaxValue, IsRef);
            ce.RegisterFunction("ISTEXT", 1, int.MaxValue, IsText);
            ce.RegisterFunction("N", 1, N);
            ce.RegisterFunction("NA", 0, NA);
            ce.RegisterFunction("TYPE", 1, Type);
        }

        private static object ErrorType(List<Expression> p)
        {
            try
            {
                p[0].Evaluate();
            }
            catch (CalcEngineException ex)
            {
                return ex switch
                {
                    NullValueException _ => 1,
                    DivisionByZeroException _ => 2,
                    CellValueException _ => 3,
                    CellReferenceException _ => 4,
                    NameNotRecognizedException _ => 5,
                    NumberException _ => 6,
                    NoValueAvailableException _ => 7,
                    _ => throw new NoValueAvailableException()
                };
            }

            throw new NoValueAvailableException();
        }

        private static object IsBlank(List<Expression> p)
        {
            var v = (string)p[0];
            var isBlank = string.IsNullOrEmpty(v);

            if (isBlank && p.Count > 1)
            {
                var sublist = p.GetRange(1, p.Count);
                isBlank = (bool)IsBlank(sublist);
            }

            return isBlank;
        }

        private static object IsErr(List<Expression> p)
        {
            try
            {
                p[0].Evaluate();
                return false;
            }
            catch (NoValueAvailableException)
            {
                return false;
            }
            catch (CalcEngineException)
            {
                return true;
            }
        }

        private static object IsError(List<Expression> p)
        {
            try
            {
                p[0].Evaluate();
                return false;
            }
            catch (CalcEngineException)
            {
                return true;
            }
        }

        private static object IsEven(List<Expression> p)
        {
            var v = p[0].Evaluate();
            if (v is double)
            {
                return Math.Abs((double)v % 2) < 1;
            }
            //TODO: Error Exceptions
            throw new ArgumentException("Expression doesn't evaluate to double");
        }

        private static object IsLogical(List<Expression> p)
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

        private static object IsNa(List<Expression> p)
        {
            try
            {
                p[0].Evaluate();
                return false;
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

        private static object IsNonText(List<Expression> p)
        {
            return !(bool)IsText(p);
        }

        private static object IsNumber(List<Expression> p)
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

        private static object IsOdd(List<Expression> p)
        {
            return !(bool)IsEven(p);
        }

        private static object IsRef(List<Expression> p)
        {
            var oe = p[0] as XObjectExpression;
            if (oe == null)
                return false;

            var crr = oe.Value as CellRangeReference;

            return crr != null;
        }

        private static object IsText(List<Expression> p)
        {
            //Evaluate Expressions
            var isText = !(bool)IsBlank(p);
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

        private static object N(List<Expression> p)
        {
            return (double)p[0];
        }

        private static object NA(List<Expression> p)
        {
            throw new NoValueAvailableException();
        }

        private static object Type(List<Expression> p)
        {
            if ((bool)IsError(p))
            {
                return 16;
            }
            if (p[0] is XObjectExpression)
            {
                // Any kind of array, ranges, names
                return 64;
            }
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

            return null;
        }
    }
}
