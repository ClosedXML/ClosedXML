// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotValueCombination : IXLPivotValueCombination
    {
        private readonly IXLPivotValue _pivotValue;

        public XLPivotValueCombination(IXLPivotValue pivotValue)
        {
            _pivotValue = pivotValue;
        }

        public IXLPivotValue And(Object item)
        {
            return _pivotValue
                .SetBaseItemValue(item)
                .SetCalculationItem(XLPivotCalculationItem.Value);
        }

        public IXLPivotValue AndNext()
        {
            return _pivotValue
                .SetCalculationItem(XLPivotCalculationItem.Next);
        }

        public IXLPivotValue AndPrevious()
        {
            return _pivotValue
                .SetCalculationItem(XLPivotCalculationItem.Previous);
        }
    }
}
