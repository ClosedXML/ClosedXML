// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueCombination
    {
        IXLPivotValue And(String item);

        IXLPivotValue AndNext();

        IXLPivotValue AndPrevious();
    }
}
