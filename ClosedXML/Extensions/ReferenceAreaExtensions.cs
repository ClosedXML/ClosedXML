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
        /// Is reference a row span (e.g. $3:7).
        /// </summary>
        public static bool IsRowSpan(this ReferenceArea reference)
        {
            return reference.First.ColumnType == ReferenceAxisType.None &&
                   reference.Second.ColumnType == ReferenceAxisType.None;
        }

        /// <summary>
        /// Is reference a column span (e.g. $B:G).
        /// </summary>
        public static bool IsColumnSpan(this ReferenceArea reference)
        {
            return reference.First.RowType == ReferenceAxisType.None &&
                   reference.Second.RowType == ReferenceAxisType.None;
        }

        /// <summary>
        /// Convert area to an absolute sheet range (regardless if the area is A1 or R1C1).
        /// </summary>
        /// <param name="area">Area to convert</param>
        /// <param name="anchor">An anchor address that is the center of R1C1 relative address.</param>
        /// <returns>Converted absolute range.</returns>
        public static XLSheetRange ToSheetRange(this ReferenceArea area, XLSheetPoint anchor)
        {
            return area.First.IsA1
                ? ToSheetRangeA1(area)
                : ToSheetRangeR1C1(area, anchor);
        }

        public static XLSheetRange ToSheetRangeA1(this ReferenceArea area)
        {
            if (area.Style != ReferenceStyle.A1)
                throw new ArgumentException(nameof(area));

            var row1 = A1ToPosition(area.First.RowType, area.First.RowValue, XLHelper.MinRowNumber);
            var col1 = A1ToPosition(area.First.ColumnType, area.First.ColumnValue, XLHelper.MinColumnNumber);
            var row2 = A1ToPosition(area.Second.RowType, area.Second.RowValue, XLHelper.MaxRowNumber);
            var col2 = A1ToPosition(area.Second.ColumnType, area.Second.ColumnValue, XLHelper.MaxColumnNumber);
            return ToSheetRange(row1, row2, col1, col2);
        }

        public static XLSheetRange ToSheetRangeR1C1(this ReferenceArea area, XLSheetPoint anchor)
        {
            if (area.Style != ReferenceStyle.R1C1)
                throw new ArgumentException(nameof(area));

            var row1 = R1C1ToPosition(area.First.RowType, area.First.RowValue, anchor.Row, XLHelper.MinRowNumber, XLHelper.MaxRowNumber);
            var col1 = R1C1ToPosition(area.First.ColumnType, area.First.ColumnValue, anchor.Column, XLHelper.MinColumnNumber, XLHelper.MaxColumnNumber);
            var row2 = R1C1ToPosition(area.Second.RowType, area.Second.RowValue, anchor.Row, XLHelper.MaxRowNumber, XLHelper.MaxRowNumber);
            var col2 = R1C1ToPosition(area.Second.ColumnType, area.Second.ColumnValue, anchor.Column, XLHelper.MaxColumnNumber, XLHelper.MaxColumnNumber);
            return ToSheetRange(row1, row2, col1, col2);
        }

        /// <summary>
        /// Shift a reference down on an area insertion. Do not shift if it would cause splits (returns <c>false</c>).
        /// </summary>
        /// <param name="reference">Reference to shift.</param>
        /// <param name="insertedArea">An area inserted into a sheet.</param>
        /// <param name="shiftedReference">The shifted reference. Can be <c>null</c>, if the reference is shifted out of sheet.</param>
        /// <returns><c>false</c> if split, <c>true</c> when shift has a rectangular reference.</returns>
        public static bool TryInsertAndShiftDown(this ReferenceArea reference, XLSheetRange insertedArea, out ReferenceArea? shiftedReference)
        {
            // Column span is never shifted
            if (reference.IsColumnSpan())
            {
                shiftedReference = reference;
                return true;
            }

            var referenceArea = reference.ToSheetRangeA1();
            if (!referenceArea.TryInsertAreaAndShiftDown(insertedArea, out var shifted))
            {
                shiftedReference = null;
                return false;
            }

            // Reference was pushed out of sheet.
            if (shifted is null)
            {
                shiftedReference = null;
                return true;
            }

            var first = Set(reference.First, shifted.Value.TopRow, referenceArea.LeftColumn);
            var second = Set(reference.Second, shifted.Value.BottomRow, referenceArea.RightColumn);
            shiftedReference = new ReferenceArea(first, second);
            return true;
        }

        /// <summary>
        /// Shift a reference to the right on an area insertion. Do not shift if it would cause splits (returns <c>false</c>).
        /// </summary>
        /// <param name="reference">Reference to shift.</param>
        /// <param name="insertedArea">An area inserted into a sheet.</param>
        /// <param name="shiftedReference">The shifted reference. Can be <c>null</c>, if the reference is shifted out of sheet.</param>
        /// <returns><c>false</c> if split, <c>true</c> when shift has a rectangular reference.</returns>
        public static bool TryInsertAndShiftRight(this ReferenceArea reference, XLSheetRange insertedArea, out ReferenceArea? shiftedReference)
        {
            // Row span is never shifted
            if (reference.IsRowSpan())
            {
                shiftedReference = reference;
                return true;
            }

            var referenceArea = reference.ToSheetRangeA1();
            if (!referenceArea.TryInsertAreaAndShiftRight(insertedArea, out var shifted))
            {
                shiftedReference = null;
                return false;
            }

            // Reference was pushed out of sheet.
            if (shifted is null)
            {
                shiftedReference = null;
                return true;
            }

            var first = Set(reference.First, referenceArea.TopRow, shifted.Value.LeftColumn);
            var second = Set(reference.Second, referenceArea.BottomRow, shifted.Value.RightColumn);
            shiftedReference = new ReferenceArea(first, second);
            return true;
        }

        public static bool TryDeleteAndShiftUp(this ReferenceArea reference, XLSheetRange deletedArea, out ReferenceArea? shiftedReference)
        {
            // Column span is never shifted
            if (reference.IsColumnSpan())
            {
                shiftedReference = reference;
                return true;
            }

            var referenceArea = reference.ToSheetRangeA1();
            if (!referenceArea.TryDeleteAreaAndShiftUp(deletedArea, out var shifted))
            {
                shiftedReference = null;
                return false;
            }

            // Reference was pushed out of sheet.
            if (shifted is null)
            {
                shiftedReference = null;
                return true;
            }

            var first = Set(reference.First, shifted.Value.TopRow, shifted.Value.LeftColumn);
            var second = Set(reference.Second, shifted.Value.BottomRow, shifted.Value.RightColumn);
            shiftedReference = new ReferenceArea(first, second);
            return true;
        }

        public static bool TryDeleteAndShiftLeft(this ReferenceArea reference, XLSheetRange deletedArea, out ReferenceArea? shiftedReference)
        {
            // Row span is never shifted
            if (reference.IsRowSpan())
            {
                shiftedReference = reference;
                return true;
            }

            var referenceArea = reference.ToSheetRangeA1();
            if (!referenceArea.TryDeleteAreaAndShiftLeft(deletedArea, out var shifted))
            {
                shiftedReference = null;
                return false;
            }

            // Reference was pushed out of sheet.
            if (shifted is null)
            {
                shiftedReference = null;
                return true;
            }

            var first = Set(reference.First, shifted.Value.TopRow, shifted.Value.LeftColumn);
            var second = Set(reference.Second, shifted.Value.BottomRow, shifted.Value.RightColumn);
            shiftedReference = new ReferenceArea(first, second);
            return true;
        }

        public static string GetDisplayStringA1(this ReferenceArea area, string? sheet)
        {
            var refA1 = area.GetDisplayStringA1();
            if (sheet is null)
                return refA1;

            var shouldEscape = NameUtils.ShouldQuote(sheet.AsSpan());
            var escapedSheetName = shouldEscape ? sheet.AlwaysEscapeSheetName() : sheet;
            return escapedSheetName + '!' + refA1;
        }

        private static XLSheetRange ToSheetRange(int row1, int row2, int col1, int col2)
        {
            // Points in the token `area` don't have to be in top left and bottom right corners,
            // e.g. D4:A1 or D1:A4. Normalize coordinates, so the sheet range has expected corners.
            var colStart = Math.Min(col1, col2);
            var colEnd = Math.Max(col1, col2);
            var rowStart = Math.Min(row1, row2);
            var rowEnd = Math.Max(row1, row2);
            return new XLSheetRange(rowStart, colStart, rowEnd, colEnd);
        }

        private static int A1ToPosition(ReferenceAxisType axisType, int position, int defaultPosition)
        {
            return axisType switch
            {
                ReferenceAxisType.Absolute => position, // $A$1 => R1C1
                ReferenceAxisType.Relative => position, // A1 => R1C1
                ReferenceAxisType.None => defaultPosition, // Only other axis specified, e.g. A:B doesn't have row.
                _ => throw new NotSupportedException()
            };
        }

        private static int R1C1ToPosition(ReferenceAxisType axisType, int position, int anchor, int defaultPosition, int dimensionSize)
        {
            switch (axisType)
            {
                case ReferenceAxisType.Absolute: // R2C5
                    return position;

                case ReferenceAxisType.Relative: // R[2]C[5]
                    {
                        var absolutePosition = anchor + position;
                        if (absolutePosition < 1)
                            return absolutePosition + dimensionSize;

                        if (absolutePosition > dimensionSize)
                            return absolutePosition - dimensionSize;

                        return absolutePosition;
                    }

                case ReferenceAxisType.None:
                    return defaultPosition; // other axis specified, e.g. R3:R5 doesn't have row.

                default:
                    throw new NotSupportedException();
            }
        }

        private static RowCol Set(RowCol rowCol, int row, int column)
        {
            var r = rowCol.RowType != ReferenceAxisType.None ? row : rowCol.RowValue;
            var c = rowCol.ColumnType != ReferenceAxisType.None ? column : rowCol.ColumnValue;
            return new RowCol(rowCol.RowType, r, rowCol.ColumnType, c, rowCol.Style);
        }
    }
}
