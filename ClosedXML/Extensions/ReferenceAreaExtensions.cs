using System;
using ClosedXML.Excel;
using ClosedXML.Parser;

namespace ClosedXML.Extensions
{
    /// <summary>
    /// Extensions method for <see cref="ReferenceArea"/>.
    /// </summary>
    internal static class ReferenceAreaExtensions
    {
        /// <summary>
        /// Convert area to an absolute sheet range (regardless if the area is A1 or R1C1).
        /// </summary>
        /// <param name="area">Area to convert</param>
        /// <param name="anchor">An anchor address that is the center of R1C1 relative address.</param>
        /// <returns>Converted absolute range.</returns>
        public static XLSheetRange ToSheetRange(this ReferenceArea area, XLSheetPoint anchor)
        {
            var firstRow = ToAbsolutePositionA1(area.First.RowType, area.First.RowValue, XLHelper.MinRowNumber);
            var firstCol = ToAbsolutePositionA1(area.First.ColumnType, area.First.ColumnValue, XLHelper.MinColumnNumber);
            var lastRow = ToAbsolutePositionA1(area.Second.RowType, area.Second.RowValue, XLHelper.MaxRowNumber);
            var lastCol = ToAbsolutePositionA1(area.Second.ColumnType, area.Second.ColumnValue, XLHelper.MaxColumnNumber);

            return new XLSheetRange(firstRow, firstCol, lastRow, lastCol);
        }

        private static int ToAbsolutePositionA1(ReferenceAxisType axisType, int position, int defaultPosition)
        {
            return axisType switch
            {
                ReferenceAxisType.Absolute => position,
                ReferenceAxisType.Relative => position,
                ReferenceAxisType.None => defaultPosition,
                _ => throw new NotSupportedException()
            };
        }
    }
}
