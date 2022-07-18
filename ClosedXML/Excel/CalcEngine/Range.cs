using OneOf;
using System;
using System.Collections.Generic;
using System.Linq;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Range is an area of cells in the workbook. It's used in formula evaluation.
    /// Every range has at least one cell.
    /// </summary>
    internal class Range
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
        /// List of areas of the range. All areas are valid and normalized. Some areas have worksheet and some don't.
        /// </summary>
        internal IReadOnlyList<XLRangeAddress> Areas { get; }

        /// <summary>
        /// An iterator over all nonblank cells of the range. Some cells can be iterated
        /// over multiple times (e.g. a union of two ranges with overlapping cells).
        /// </summary>
        public IEnumerable<ScalarValue> GetCellsValues(CalcContext ctx)
        {
            // TODO: Optimize to return only nonblank through CellCollection
            foreach (var area in Areas)
            {
                for (var row = area.FirstAddress.RowNumber; row <= area.LastAddress.RowNumber; ++row)
                {
                    for (var column = area.FirstAddress.ColumnNumber; column <= area.LastAddress.ColumnNumber; ++column)
                    {
                        var cellValue = ctx.GetCellValueOrBlank(area.Worksheet, row, column);
                        if (cellValue is not null)
                        {
                            yield return cellValue.Value;
                        }
                    }
                }
            }
        }

        public static OneOf<Range, Error1> RangeOp(Range lhs, Range rhs)
        {
            var sheets = lhs.Areas.Select(a => a.Worksheet).Concat(rhs.Areas.Select(a => a.Worksheet))
                .Where(a => a.Worksheet is not null).Distinct().ToList();
            if (sheets.Count > 1)
            {
                return Error1.Value;
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

        public static OneOf<Range, Error1> Intersect(Range lhs, Range rhs, CalcContext ctx)
        {
            var sheets = lhs.Areas.Select(a => a.Worksheet ?? ctx.Worksheet)
                .Concat(rhs.Areas.Select(a => a.Worksheet ?? ctx.Worksheet))
                .Distinct().ToList();
            if (sheets.Count != 1)
                return Error1.Value;

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
                    intersections.Add((XLRangeAddress)intersectedArea);
            }

            return intersections.Any() ? new Range(intersections) : Error1.Null;
        }

        public OneOf<ScalarValue, Error1> ImplicitIntersection(CalcContext ctx)
        {
            if (Areas.Count != 1)
                return Error1.Value;

            var area = Areas.Single();
            if (area.RowSpan == 1 && area.ColumnSpan == 1)
                return ctx.GetCellValue(area.Worksheet, area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber);

            var column = ctx.FormulaAddress.ColumnNumber;
            var row = ctx.FormulaAddress.RowNumber;

            if (area.ColumnSpan == 1 && area.FirstAddress.RowNumber <= row && row <= area.LastAddress.RowNumber)
                return ctx.GetCellValue(area.Worksheet, row, area.FirstAddress.ColumnNumber);

            if (area.RowSpan == 1 && area.FirstAddress.ColumnNumber <= column && column <= area.LastAddress.ColumnNumber)
                return ctx.GetCellValue(area.Worksheet, area.FirstAddress.RowNumber, column);

            return Error1.Value;
        }
    }
}
