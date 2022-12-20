using System;
using System.Collections.Generic;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Logical
    {
        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("AND", 1, int.MaxValue, And, AllowRange.All);
            ce.RegisterFunction("FALSE", 0, 0, Adapt(False), FunctionFlags.Scalar);
            ce.RegisterFunction("IF", 2, 3, If);
            ce.RegisterFunction("IFERROR",2,IfError);
            ce.RegisterFunction("NOT", 1, 1, AdaptCoerced(Not), FunctionFlags.Scalar);
            ce.RegisterFunction("OR", 1, int.MaxValue, Or);
            ce.RegisterFunction("TRUE", 0, 0, Adapt(True), FunctionFlags.Scalar);
        }

        private static object And(List<Expression> p)
        {
            var b = true;
            foreach (var v in p)
            {
                b = b && v;
            }
            return b;
        }

        private static object Or(List<Expression> p)
        {
            var b = false;
            foreach (var v in p)
            {
                b = b || v;
            }
            return b;
        }

        private static AnyValue Not(Boolean value)
        {
            return !value;
        }

        private static object If(List<Expression> p)
        {
            if (p[0])
            {
                return p[1].Evaluate();
            }
            else if (p.Count > 2)
            {
                if (p[2] is EmptyValueExpression)
                    return false;
                else
                    return p[2].Evaluate();
            }
            else return false;
        }

        private static AnyValue True()
        {
            return true;
        }

        private static AnyValue False()
        {
            return false;
        }

        private static object IfError(List<Expression> p)
        {
            try
            {
                var value = p[0].Evaluate();
                if (value is XLError)
                    return p[1].Evaluate();

                return value;
            }
            catch (ArgumentException)
            {
                return p[1].Evaluate();
            }
        }
    }
}
