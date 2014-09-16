using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Is
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("ISBLANK", 1, IsBlank);
        }

        static object IsBlank(List<Expression> p)
        {
            var v = (string)p[0];
            return String.IsNullOrEmpty(v);
        }
    }
}
