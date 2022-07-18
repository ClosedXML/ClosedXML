using OneOf;
using System;
using System.Collections.Generic;
using System.Linq;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Range>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Range is an area of cells in the workbook. It's used in formula evaluation.
    /// Every range has at least one cell.
    /// </summary>
    internal class Range : IEnumerable<AnyValue>
    {
        public Range(XLRangeAddress area)
        {
            if (!area.IsNormalized)
                throw new ArgumentException();

            Areas = new List<XLRangeAddress>(1) { area };
        }

        private Range(List<XLRangeAddress> areas)
        {
            Areas = areas;
        }

        /// <summary>
        /// List of areas of the range. All areas are normalized. Some areas have worksheet and some don't.
        /// </summary>
        internal IReadOnlyList<XLRangeAddress> Areas { get; }

        /// <summary>
        /// An iterator over all nonblank cells of the range. Some cells can be iterated
        /// over multiple times (e.g. a union of two ranges with overlapping cells).
        /// </summary>
        public IEnumerator<AnyValue> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

        public static OneOf<Range, Error1> RangeOp(Range lhs, Range rhs)
        {
            var sheets = lhs.Areas.Select(a => a.Worksheet).Concat(rhs.Areas.Select(a => a.Worksheet)).Distinct().ToList();
            if (sheets.Count != 1)
            {
                return Error1.CellValue;
            }

            var minCol = XLHelper.MaxColumnNumber;
            var maxCol = 1;
            var minRow = XLHelper.MaxRowNumber;
            var maxRow = 1;
            foreach (var area in lhs.Areas.Concat(rhs.Areas))
            {
                // Areas are normalized, so I don't have to check opposite corners
                minRow = Math.Min(minRow, area.FirstAddress.RowNumber);
                maxRow = Math.Max(maxRow, area.LastAddress.RowNumber);
                minCol = Math.Min(minCol, area.FirstAddress.ColumnNumber);
                maxCol = Math.Max(maxCol, area.LastAddress.ColumnNumber);
            }

            var sheet = sheets.Single();
            return new Range(new XLRangeAddress(new XLAddress(sheet, minRow, minCol, true, true), new XLAddress(sheet, maxRow, maxCol, true, true)));
        }

        public static Range UnionOp(Range lhs, Range rhs)
        {
            return new Range(lhs.Areas.Concat(rhs.Areas).ToList());
        }

        public static Range Intersection(Range lhs, Range rhs)
        {
            throw new NotImplementedException();
        }
    }
}
