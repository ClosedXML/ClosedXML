using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel.CalcEngine
{
    static class Logical
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("AND", 1, int.MaxValue, And);
            ce.RegisterFunction("OR", 1, int.MaxValue, Or);
            ce.RegisterFunction("NOT", 1, Not);
            ce.RegisterFunction("IF", 3, If);
            ce.RegisterFunction("TRUE", 0, True);
            ce.RegisterFunction("FALSE", 0, False);
        }

        static object And(List<Expression> p)
        {
            var b = true;
            foreach (var v in p)
            {
                b = b && (bool)v;
            }
            return b;
        }
        static object Or(List<Expression> p)
        {
            var b = false;
            foreach (var v in p)
            {
                b = b || (bool)v;
            }
            return b;
        }
        static object Not(List<Expression> p)
        {
            return !(bool)p[0];
        }
        static object If(List<Expression> p)
        {
            return (bool)p[0] 
                ? p[1].Evaluate() 
                : p[2].Evaluate();
        }
        static object True(List<Expression> p)
        {
            return true;
        }
        static object False(List<Expression> p)
        {
            return false;
        }
    }
}
