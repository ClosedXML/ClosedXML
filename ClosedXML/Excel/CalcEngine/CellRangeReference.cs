using System;
using System.Collections;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CellRangeReference : IValueObject, IEnumerable
    {
        private IXLRange _range;
        private XLCalcEngine _ce;

        public CellRangeReference(IXLRange range, XLCalcEngine ce)
        {
            _range = range;
            _ce = ce;
        }

        public IXLRange Range { get { return _range; } }

        // ** IValueObject
        public object GetValue()
        {
            return GetValue(_range.FirstCell());
        }

        // ** IEnumerable
        public IEnumerator GetEnumerator()
        {
            var maxRow = Math.Min(_range.RangeAddress.LastAddress.RowNumber, _range.Worksheet.LastCellUsed().Address.RowNumber);
            var maxCol = Math.Min(_range.RangeAddress.LastAddress.ColumnNumber, _range.Worksheet.LastCellUsed().Address.ColumnNumber);
            using (var trimmedRange = (XLRangeBase)_range.Worksheet.Range(_range.FirstCell().Address, new XLAddress(maxRow, maxCol, false, false)))
                return trimmedRange.CellValues().GetEnumerator();
        }

        private Boolean _evaluating;

        // ** implementation
        private object GetValue(IXLCell cell)
        {
            if (_evaluating || (cell as XLCell).IsEvaluating)
            {
                throw new InvalidOperationException($"Circular Reference occured during evaluation. Cell: {cell.Address.ToString(XLReferenceStyle.Default, true)}");
            }
            try
            {
                _evaluating = true;
                var f = cell.FormulaA1;
                if (String.IsNullOrWhiteSpace(f))
                    return cell.Value;
                else
                {
                    return (cell as XLCell).Evaluate();
                }
            }
            finally
            {
                _evaluating = false;
            }
        }
    }
}
