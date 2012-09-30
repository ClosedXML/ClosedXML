using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.CalcEngine.Functions.Subtotals
{
    abstract class Subtotal
    {
        protected readonly List<Expression> exprList;

        protected Subtotal(List<Expression> list)
        {
            exprList = list;
        }
        public abstract Object Evaluate();

        public static Object GetSubtotal(Int32 fId, List<Expression> list)
        {

            switch (fId)
            {
                case 1:
                    return Average.GetSubtotal(list);
                default:
                    throw new ArgumentException("Function not supported.");
            }
        }
    }
}
