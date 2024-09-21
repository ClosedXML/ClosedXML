#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using ClosedXML.Excel.CalcEngine.Exceptions;

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

            foreach (var c in CellValues(trimmedRange))
                yield return c.ToObject();
        }

        private Boolean _evaluating;

        // ** implementation
        private object GetValue(IXLCell cell)
        {
            if (_evaluating)
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

        internal IEnumerable<XLCellValue> CellValues() => CellValues(Range);

        private static IEnumerable<XLCellValue> CellValues(IXLRangeBase range)
        {
            var sheet = (XLWorksheet)range.Worksheet;
            for (int ro = range.RangeAddress.FirstAddress.RowNumber; ro <= range.RangeAddress.LastAddress.RowNumber; ro++)
            {
                for (int co = range.RangeAddress.FirstAddress.ColumnNumber; co <= range.RangeAddress.LastAddress.ColumnNumber; co++)
                {
                    var value = GetCellValue(sheet, new XLSheetPoint(ro, co));
                    yield return value;
                }
            }
        }

        private static XLCellValue GetCellValue(XLWorksheet sheet, XLSheetPoint point)
        {
            var cell = sheet.GetCell(point);
            if (cell is null)
                return Blank.Value;

            if (cell.Formula is null || !cell.Formula.IsDirty)
                return cell.CachedValue;

            throw new GettingDataException(new XLBookPoint(sheet.SheetId, point));
        }
    }
}
