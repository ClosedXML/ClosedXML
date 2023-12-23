using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A comparer of rows in a range. It uses semantic of a sort feature in Excel.
    /// </summary>
    /// <remarks>
    /// The comparer should work separate from data, but it would necessitate to sort over
    /// <see cref="XLRangeRow"/>. That would require to not only instantiate a new object for each
    /// sorted row, but since <see cref="XLRangeRow"/>, it would also be be tracked in range
    /// repository, slowing each subsequent operation. To improve performance, comparer has
    /// reference to underlaying data and compares row numbers that can be stores in a single
    /// allocated array of indexes.
    /// </remarks>
    internal class XLRangeRowsSortComparer : IComparer<int>
    {
        private readonly List<(int ColumnNumber, XLCellValueSortComparer Comparer)> _columnComparers;
        private readonly ValueSlice _valueSlice;

        internal XLRangeRowsSortComparer(XLWorksheet sheet, XLSheetRange sortRange, IXLSortElements sortColumns)
        {
            if (!sortColumns.Any())
                throw new ArgumentException("Empty sort specification.");

            if (sortRange.Width < sortColumns.Max(x => x.ElementNumber))
                throw new ArgumentException("Range has fewer columns that sort specification.");

            _valueSlice = sheet.Internals.CellsCollection.ValueSlice;
            _columnComparers = sortColumns.Select(se => (se.ElementNumber + sortRange.LeftColumn - 1, new XLCellValueSortComparer(se))).ToList();
        }

        public int Compare(int rowNumber1, int rowNumber2)
        {
            foreach (var (columnNumber, comparer) in _columnComparers)
            {
                var row1 = _valueSlice.GetCellValue(new XLSheetPoint(rowNumber1, columnNumber));
                var row2 = _valueSlice.GetCellValue(new XLSheetPoint(rowNumber2, columnNumber));
                var comparison = comparer.Compare(row1, row2);
                if (comparison != 0)
                    return comparison;
            }

            // Row sort should be stable, because otherwise we could randomly switch cells
            // with different formats on subsequent sorts. BCL doesn't support in-place
            // stable sort (Array/List.Sort) directly, only LINQ does it (thus extra copy).
            // Note that stable sort has worse worst case O(N*log(N)^2).
            //
            // As a workaround for stable sort, if all values look same, use the order of rows.
            return rowNumber1 - rowNumber2;
        }
    }
}
