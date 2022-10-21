using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Reference is a collection of cells in the workbook. It's used in formula evaluation.
    /// Every reference has at least one cell.
    /// </summary>
    internal class Reference
    {
        public Reference(XLRangeAddress area)
        {
            if (!area.IsNormalized)
                throw new ArgumentException();

            Areas = new List<XLRangeAddress>(1) { area };
        }

        /// <summary>
        /// Ctor that reuses parameter to keep allocations low - don't modify the collection after it is passed to ctor.
        /// </summary>
        public Reference(List<XLRangeAddress> areas)
        {
            if (areas.Count < 1)
                throw new ArgumentException();

            Areas = areas;
        }

        public Reference(IXLRanges ranges)
        {
            if (ranges.Count < 1)
                throw new ArgumentException();

            var areas = new List<XLRangeAddress>(ranges.Count);
            foreach (var range in ranges)
                areas.Add((XLRangeAddress)range.RangeAddress);

            Areas = areas;
        }

        /// <summary>
        /// List of areas of the range (at least one). All areas are valid and normalized. Some areas have worksheet and some don't.
        /// </summary>
        internal IReadOnlyList<XLRangeAddress> Areas { get; }

        /// <summary>
        /// An iterator over all nonblank cells of the range. Some cells can be iterated
        /// over multiple times (e.g. a union of two ranges with overlapping cells).
        /// </summary>
        public IEnumerable<ScalarValue> GetCellsValues(CalcContext ctx)
        {
            foreach (var area in Areas)
            {
                for (var row = area.FirstAddress.RowNumber; row <= area.LastAddress.RowNumber; ++row)
                {
                    for (var column = area.FirstAddress.ColumnNumber; column <= area.LastAddress.ColumnNumber; ++column)
                    {
                        var cellValue = ctx.GetCellValue(area.Worksheet, row, column);
                        if (!cellValue.IsBlank)
                        {
                            yield return cellValue;
                        }
                    }
                }
            }
        }

        public static OneOf<Reference, XLError> RangeOp(Reference lhs, Reference rhs, XLWorksheet contextWorksheet)
        {
            var lhsWorksheets = lhs.Areas.Count == 1
                ? lhs.Areas.Select(a => a.Worksheet).Where(ws => ws is not null).ToList()
                : lhs.Areas.Select(a => a.Worksheet ?? contextWorksheet).Where(ws => ws is not null).Distinct().ToList();
            if (lhsWorksheets.Count() > 1)
                return XLError.IncompatibleValue;

            var lhsWorksheet = lhsWorksheets.SingleOrDefault();

            var rhsWorksheets = rhs.Areas.Count == 1
                ? rhs.Areas.Select(a => a.Worksheet).Where(ws => ws is not null).ToList()
                : rhs.Areas.Select(a => a.Worksheet ?? contextWorksheet).Where(ws => ws is not null).Distinct().ToList();
            if (rhsWorksheets.Count() > 1)
                return XLError.IncompatibleValue;

            var rhsWorksheet = rhsWorksheets.SingleOrDefault();

            if (rhsWorksheet is not null)
            {
                if ((lhsWorksheet ?? contextWorksheet) != rhsWorksheet)
                    return XLError.IncompatibleValue;
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

            var sheet = lhsWorksheet ?? rhsWorksheet;
            return new Reference(new XLRangeAddress(
                new XLAddress(sheet, minRow, minCol, false, false),
                new XLAddress(sheet, maxRow, maxCol, false, false)));
        }

        public static Reference UnionOp(Reference lhs, Reference rhs)
        {
            return new Reference(lhs.Areas.Concat(rhs.Areas).ToList());
        }

        public static OneOf<Reference, XLError> Intersect(Reference lhs, Reference rhs, CalcContext ctx)
        {
            var sheets = lhs.Areas.Select(a => a.Worksheet ?? ctx.Worksheet)
                .Concat(rhs.Areas.Select(a => a.Worksheet ?? ctx.Worksheet))
                .Distinct().ToList();
            if (sheets.Count != 1)
                return XLError.IncompatibleValue;

            var sheet = sheets.Single();
            var intersections = new List<XLRangeAddress>();
            foreach (var leftArea in lhs.Areas)
            {
                var intersectedArea = leftArea.WithWorksheet(sheet);
                foreach (var rightArea in rhs.Areas)
                {
                    intersectedArea = intersectedArea.Intersection(rightArea.WithWorksheet(sheet));
                    if (!intersectedArea.IsValid)
                        break;
                }

                if (intersectedArea.IsValid)
                    intersections.Add(intersectedArea);
            }

            return intersections.Any() ? new Reference(intersections) : XLError.NullValue;
        }

        /// <summary>
        /// Do an implicit intersection of an address.
        /// </summary>
        /// <param name="formulaAddress"></param>
        /// <returns>An address of the intersection or error if intersection failed.</returns>
        public OneOf<Reference, XLError> ImplicitIntersection(IXLAddress formulaAddress)
        {
            if (Areas.Count != 1)
                return XLError.IncompatibleValue;

            var area = Areas.Single();
            if (area.RowSpan == 1 && area.ColumnSpan == 1)
                return this;

            var column = formulaAddress.ColumnNumber;
            var row = formulaAddress.RowNumber;

            if (area.ColumnSpan == 1 && area.FirstAddress.RowNumber <= row && row <= area.LastAddress.RowNumber)
            {
                var intersection = new XLAddress(area.Worksheet, row, area.FirstAddress.ColumnNumber, false, false);
                return new Reference(new XLRangeAddress(intersection, intersection));
            }

            if (area.RowSpan == 1 && area.FirstAddress.ColumnNumber <= column && column <= area.LastAddress.ColumnNumber)
            {
                var intersection = new XLAddress(area.Worksheet, area.FirstAddress.RowNumber, column, false, false);
                return new Reference(new XLRangeAddress(intersection, intersection));
            }

            return XLError.IncompatibleValue;
        }

        internal bool IsSingleCell()
        {
            return Areas.Count == 1 && Areas[0].IsSingleCell();
        }

        internal bool TryGetSingleCellValue(out ScalarValue value, CalcContext ctx)
        {
            if (!IsSingleCell())
            {
                value = default;
                return false;
            }

            var area = Areas.Single();
            value = ctx.GetCellValue(area.Worksheet, area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber);
            return true;
        }

        internal OneOf<Array, XLError> ToArray(CalcContext context)
        {
            if (Areas.Count != 1)
                throw new NotImplementedException();

            var area = Areas.Single();

            return new ReferenceArray(area, context);
        }

        public OneOf<Array, XLError> Apply(Func<ScalarValue, ScalarValue> op, CalcContext context)
        {
            if (Areas.Count != 1)
                return XLError.IncompatibleValue;

            var area = Areas.Single();
            var width = area.ColumnSpan;
            var height = area.RowSpan;
            var startColumn = area.FirstAddress.ColumnNumber;
            var startRow = area.FirstAddress.RowNumber;
            var data = new ScalarValue[height, width];
            for (int y = 0; y < height; ++y)
            {
                for (int x = 0; x < width; ++x)
                {
                    var row = startRow + y;
                    var column = startColumn + x;
                    var cellValue = context.GetCellValue(area.Worksheet, row, column);
                    data[y, x] = op(cellValue);
                }
            }

            return new ConstArray(data);
        }
    }
}
