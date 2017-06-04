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
            ce.RegisterFunction("ERRORTYPE",1,ErrorType); 
            ce.RegisterFunction("ISBLANK", 1,int.MaxValue, IsBlank);
            ce.RegisterFunction("ISERR",1, int.MaxValue, IsErr);
            ce.RegisterFunction("ISERROR",1, int.MaxValue, IsError);
            ce.RegisterFunction("ISEVEN",1, IsEven);
            ce.RegisterFunction("ISLOGICAL",1,int.MaxValue,IsLogical);
            ce.RegisterFunction("ISNA",1, int.MaxValue, IsNa);
            ce.RegisterFunction("ISNONTEXT",1, int.MaxValue, IsNonText);
            ce.RegisterFunction("ISNUMBER",1, int.MaxValue, IsNumber);
            ce.RegisterFunction("ISODD",1,IsOdd);
            ce.RegisterFunction("ISREF",1, int.MaxValue, IsRef);
            ce.RegisterFunction("ISTEXT", 1, int.MaxValue, IsText);
            ce.RegisterFunction("N",1,N);
            ce.RegisterFunction("NA",0,NA);
            ce.RegisterFunction("TYPE",1,Type);
        }

        static object ErrorType(List<Expression> p)
        {
            //TODO: Write Code
            throw new NotSupportedException();;
        }

        static object IsBlank(List<Expression> p)
        {
            var v = (string) p[0];
            var isBlank = string.IsNullOrEmpty(v);


            if (isBlank && p.Count > 1) {
                var sublist = p.GetRange(1, p.Count);
                isBlank = (bool)IsBlank(sublist);
            }

            return isBlank;
        }

        //TODO: Support for Error Values
        static object IsErr(List<Expression> p)
        {
            //TODO: Write Code
            throw new NotSupportedException();
        }
            
        static object IsError(List<Expression> p)
        {
            //TODO: Write Code
            throw new NotSupportedException();
        }

        static object IsEven(List<Expression> p)
        {
            var v = p[0].Evaluate();
            if (v is double)
            {
                return Math.Abs((double) v%2) < 1;
            }
            //TODO: Error Exceptions
            throw new ArgumentException("Expression doesn't evaluate to double");
        }

        static object IsLogical(List<Expression> p)
        {
            var v = p[0].Evaluate();
            var isLogical = v is bool;
            
            if (isLogical && p.Count > 1)
            {
                var sublist = p.GetRange(1, p.Count);
                isLogical = (bool) IsLogical(sublist);
            }

            return isLogical;
        }

        static object IsNa(List<Expression> p)
        {
            //TODO: Write Code
            throw new NotSupportedException();;
        }

        static object IsNonText(List<Expression> p)
        {
            return !(bool) IsText(p);
        }

        static object IsNumber(List<Expression> p)
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
                    var stringValue = (string) v;
                    double.Parse(stringValue.TrimEnd('%', ' '), NumberStyles.Any);
                    isNumber = true;
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

        static object IsOdd(List<Expression> p)
        {
            return !(bool) IsEven(p);
        }

        static object IsRef(List<Expression> p)
        {
            //TODO: Write Code
            throw new NotSupportedException();;
        }

        static object IsText(List<Expression> p)
        {
            //Evaluate Expressions
            var isText = !(bool) IsBlank(p);
            if (isText)
            {
                isText = !(bool) IsNumber(p);
            }
            if (isText)
            {
                isText = !(bool) IsLogical(p);
            }
            return isText;
        }

        static object N(List<Expression> p)
        {
            return (double) p[0];
        }

        static object NA(List<Expression> p)
        {
            //TODO: Write Code
            throw new NotSupportedException();;
        }

        static object Type(List<Expression> p)
        {
            if ((bool) IsNumber(p))
            {
                return 1;
            }
            if ((bool) IsText(p))
            {
                return 2;
            }
            if ((bool) IsLogical(p))
            {
                return 4;
            }
            if ((bool) IsError(p))
            {
                return 16;
            }
            if(p.Count > 1)
            {
                return 64;
            }
            return null;
        }
    }
}