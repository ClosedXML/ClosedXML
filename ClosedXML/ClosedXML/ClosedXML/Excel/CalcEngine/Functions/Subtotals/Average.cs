using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.CalcEngine.Functions.Subtotals
{
    class Average
    {
        public static Object GetSubtotal(List<Expression> list)
        {
            var tally = new Tally();
            foreach (var e in list)
            {
                tally.Add(e);
            }
            return tally.Average();
        }
    }
}
