using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeColumnsSortComparer : IComparer<int>
    {
        private readonly List<(int RowNumber, XLCellValueSortComparer Comparer)> _rowComparers;
        private readonly ValueSlice _valueSlice;

        internal XLRangeColumnsSortComparer(XLWorksheet sheet, XLSheetRange sortRange, IXLSortElements sortRows)
        {
            if (!sortRows.Any())
                throw new ArgumentException("Empty sort specification.");

            if (sortRange.Width < sortRows.Max(x => x.ElementNumber))
                throw new ArgumentException("Range has fewer columns that sort specification.");

            _valueSlice = sheet.Internals.CellsCollection.ValueSlice;
            _rowComparers = sortRows.Select(se => (se.ElementNumber + sortRange.TopRow - 1, new XLCellValueSortComparer(se))).ToList();
        }

        public int Compare(int colNumber1, int colNumber2)
        {
            foreach (var (rowNumber, comparer) in _rowComparers)
            {
                var col1 = _valueSlice.GetCellValue(new XLSheetPoint(rowNumber, colNumber1));
                var col2 = _valueSlice.GetCellValue(new XLSheetPoint(rowNumber, colNumber2));
                var comparison = comparer.Compare(col1, col2);
                if (comparison != 0)
                    return comparison;
            }

            // Workaround for stable sort, see XLRangeRowsSortComparer.
            return colNumber1 - colNumber2;
        }
    }
}
