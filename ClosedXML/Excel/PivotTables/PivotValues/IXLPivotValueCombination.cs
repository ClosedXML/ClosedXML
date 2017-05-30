using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueCombination
    {
        IXLPivotValue And(String item);
        IXLPivotValue AndPrevious();
        IXLPivotValue AndNext();
    }
}
