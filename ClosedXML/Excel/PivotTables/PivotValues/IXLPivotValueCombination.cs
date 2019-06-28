// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueCombination
    {
        IXLPivotValue And(Object item);

        IXLPivotValue AndNext();

        IXLPivotValue AndPrevious();
    }
}
