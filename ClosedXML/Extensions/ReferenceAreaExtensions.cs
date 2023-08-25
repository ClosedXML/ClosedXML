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
            var firstRow = ToAbsolutePosition(area.First.RowType, area.First.RowValue, anchor.Row, XLHelper.MinRowNumber);
            var firstCol = ToAbsolutePosition(area.First.ColumnType, area.First.ColumnValue, anchor.Column, XLHelper.MinColumnNumber);
            var lastRow = ToAbsolutePosition(area.Second.RowType, area.Second.RowValue, anchor.Row, XLHelper.MaxRowNumber);
            var lastCol = ToAbsolutePosition(area.Second.ColumnType, area.Second.ColumnValue, anchor.Column, XLHelper.MaxColumnNumber);

            return new XLSheetRange(firstRow, firstCol, lastRow, lastCol);
        }

        private static int ToAbsolutePosition(ReferenceAxisType axisType, int position, int anchorPosition, int defaultPosition)
        {
            return axisType switch
            {
                ReferenceAxisType.Absolute => position,
                ReferenceAxisType.Relative => anchorPosition + position,
                ReferenceAxisType.None => defaultPosition,
                _ => throw new NotSupportedException()
            };
        }
    }
}
