using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Logical
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("AND", 1, int.MaxValue, And);
            ce.RegisterFunction("OR", 1, int.MaxValue, Or);
            ce.RegisterFunction("NOT", 1, Not);
            ce.RegisterFunction("IF", 2, 3, If);
            ce.RegisterFunction("TRUE", 0, True);
            ce.RegisterFunction("FALSE", 0, False);
            ce.RegisterFunction("IFERROR",2,IfError);
        }

        static object And(List<Expression> p)
        {
            var b = true;
            foreach (var v in p)
            {
                b = b && v;
            }
            return b;
        }

        static object Or(List<Expression> p)
        {
            var b = false;
            foreach (var v in p)
            {
                b = b || v;
            }
            return b;
        }

        static object Not(List<Expression> p)
        {
            return !p[0];
        }

        static object If(List<Expression> p)
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

        static object True(List<Expression> p)
        {
            return true;
        }

        static object False(List<Expression> p)
        {
            return false;
        }

        static object IfError(List<Expression> p)
        {
            try
            {
                return p[0].Evaluate();
            }
            catch (ArgumentException)
            {
                return p[1].Evaluate();
            }
        }
    }
}
