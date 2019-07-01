using System;
using System.Collections;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CellRangeReference : IValueObject, IEnumerable
    {
        private readonly XLCalcEngine _ce;

        public CellRangeReference(IXLRange range, XLCalcEngine ce)
        {
            Range = range;
            _ce = ce;
        }

        public IXLRange Range { get; }

        // ** IValueObject
        public object GetValue()
        {
            return GetValue(Range.FirstCell());
        }

        // ** IEnumerable
        public IEnumerator GetEnumerator()
        {
            var lastCellAddress = Range.Worksheet.LastCellUsed().Address;
            var maxRow = Math.Min(Range.RangeAddress.LastAddress.RowNumber, lastCellAddress.RowNumber);
            var maxColumn = Math.Min(Range.RangeAddress.LastAddress.ColumnNumber, lastCellAddress.ColumnNumber);
            var trimmedRange = (XLRangeBase)Range.Worksheet.Range(Range.FirstCell().Address,
                new XLAddress(maxRow, maxColumn, fixedRow: false, fixedColumn: false));
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
                    return (cell as XLCell).EvaluateInternal();
                }
            }
            finally
            {
                _evaluating = false;
            }
        }
    }
}
