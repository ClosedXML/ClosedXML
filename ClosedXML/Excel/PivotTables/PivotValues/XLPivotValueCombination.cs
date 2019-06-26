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

        public IXLPivotValue And(String item)
        {
            _pivotValue.BaseItem = item;
            _pivotValue.CalculationItem = XLPivotCalculationItem.Value;
            return _pivotValue;
        }

        public IXLPivotValue AndNext()
        {
            _pivotValue.CalculationItem = XLPivotCalculationItem.Next;
            return _pivotValue;
        }

        public IXLPivotValue AndPrevious()
        {
            _pivotValue.CalculationItem = XLPivotCalculationItem.Previous;
            return _pivotValue;
        }
    }
}
