#nullable disable

using System;
using System.Collections;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CellRangeReference : IValueObject, IEnumerable
    {
        public CellRangeReference(IXLRange range)
        {
            Range = range;
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
            if (Range.Worksheet.IsEmpty(XLCellsUsedOptions.AllContents))
                yield break;

            var lastCellAddress = Range.Worksheet.LastCellUsed().Address;
            var maxRow = Math.Min(Range.RangeAddress.LastAddress.RowNumber, lastCellAddress.RowNumber);
            var maxColumn = Math.Min(Range.RangeAddress.LastAddress.ColumnNumber, lastCellAddress.ColumnNumber);

            var trimmedRange = (XLRangeBase)Range.Worksheet
                .Range(
                    Range.FirstCell().Address,
                    new XLAddress(maxRow, maxColumn, fixedRow: false, fixedColumn: false)
                );

            foreach (var c in trimmedRange.CellValues())
                yield return c.ToObject();
        }

        private Boolean _evaluating;

        // ** implementation
        private object GetValue(IXLCell cell)
        {
            if (_evaluating || ((XLCell)cell).IsEvaluating)
            {
                throw new InvalidOperationException($"Circular Reference occurred during evaluation. Cell: {cell.Address.ToString(XLReferenceStyle.Default, true)}");
            }
            try
            {
                _evaluating = true;
                return cell.Value.ToObject();
            }
            finally
            {
                _evaluating = false;
            }
        }
    }
}
