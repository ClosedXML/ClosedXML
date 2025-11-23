using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A representation of a <c>ST_Ref</c>, i.e. an area in a sheet (no reference to the sheet).
    /// </summary>
    internal readonly struct XLSheetRange : IEquatable<XLSheetRange>, IEnumerable<XLSheetPoint>
    {
        internal XLSheetRange(XLSheetPoint point)
            : this(point, point)
        {
        }

        internal XLSheetRange(XLSheetPoint firstPoint, XLSheetPoint lastPoint)
        {
            FirstPoint = firstPoint;
            LastPoint = lastPoint;
        }

        public XLSheetRange(Int32 rowStart, Int32 columnStart, Int32 rowEnd, Int32 columnEnd)
            : this(new XLSheetPoint(rowStart, columnStart), new XLSheetPoint(rowEnd, columnEnd))
        {
        }

        /// <summary>
        /// A range that covers whole worksheet.
        /// </summary>
        public static readonly XLSheetRange Full = new(
            new XLSheetPoint(XLHelper.MinRowNumber, XLHelper.MinColumnNumber),
            new XLSheetPoint(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber));

        /// <summary>
        /// Top-left point of the sheet range.
        /// </summary>
        public readonly XLSheetPoint FirstPoint;

        /// <summary>
        /// Bottom-right point of the sheet range.
        /// </summary>
        public readonly XLSheetPoint LastPoint;

        public int Width => LastPoint.Column - FirstPoint.Column + 1;

        public int Height => LastPoint.Row - FirstPoint.Row + 1;

        /// <summary>
        /// The left column number of the range. From 1 to <see cref="XLHelper.MaxColumnNumber"/>.
        /// </summary>
        public int LeftColumn => FirstPoint.Column;

        /// <summary>
        /// The right column number of the range. From 1 to <see cref="XLHelper.MaxColumnNumber"/>.
        /// Greater or equal to <see cref="LeftColumn"/>.
        /// </summary>
        public int RightColumn => LastPoint.Column;

        /// <summary>
        /// The top row number of the range. From 1 to <see cref="XLHelper.MaxRowNumber"/>.
        /// </summary>
        public int TopRow => FirstPoint.Row;

        /// <summary>
        /// The bottom row number of the range. From 1 to <see cref="XLHelper.MaxRowNumber"/>.
        /// Greater or equal to <see cref="TopRow"/>.
        /// </summary>
        public int BottomRow => LastPoint.Row;

        /// <summary>
        /// Does area span from first to last column?
        /// </summary>
        internal bool HasFullRowWidth => LeftColumn == XLHelper.MinColumnNumber && RightColumn == XLHelper.MaxColumnNumber;

        /// <summary>
        /// Does area span from first to last row?
        /// </summary>
        internal bool HasFullColumnHeight => TopRow == XLHelper.MinRowNumber && BottomRow == XLHelper.MaxRowNumber;

        public override bool Equals(object? obj)
        {
            return obj is XLSheetRange range && Equals(range);
        }

        public bool Equals(XLSheetRange other)
        {
            return FirstPoint.Equals(other.FirstPoint) && LastPoint.Equals(other.LastPoint);
        }

        public override int GetHashCode()
        {
            return FirstPoint.GetHashCode() ^ LastPoint.GetHashCode();
        }

        public static bool operator ==(XLSheetRange left, XLSheetRange right) => left.Equals(right);

        public static bool operator !=(XLSheetRange left, XLSheetRange right) => !(left == right);


        /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
        public static XLSheetRange Parse(String input) => Parse(input.AsSpan());

        /// <summary>
        /// Parse point per type <c>ST_Ref</c> from
        /// <a href="https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e7f22870-88a1-4c06-8e5f-d035b1179c50">2.1.1119 Part 4 Section 3.18.64, ST_Ref (Cell Range Reference)</a>
        /// </summary>
        /// <remarks>Can be one cell reference (A1) or two separated by a colon (A1:B2). First reference is always in top left corner</remarks>
        /// <param name="input">Input text</param>
        /// <exception cref="FormatException">If the input doesn't match expected grammar.</exception>
        public static XLSheetRange Parse(ReadOnlySpan<char> input)
        {
            if (!TryParse(input, out var area))
                throw new FormatException($"Area reference doesn't have correct format: '{input.ToString()}'.");

            return area;
        }

        /// <summary>
        /// Try to parse area. Doesn't accept any extra whitespace anywhere in the input. Letters
        /// must be upper case. Area can specify one corner (<c>A1</c>) or both corners (<c>A1:B3</c>).
        /// </summary>
        public static bool TryParse(ReadOnlySpan<char> input, out XLSheetRange area)
        {
            var separatorIndex = input.IndexOf(':');
            if (separatorIndex == -1)
            {
                if (!XLSheetPoint.TryParse(input, out var sheetPoint))
                {
                    area = default;
                    return false;
                }

                area = new XLSheetRange(sheetPoint, sheetPoint);
                return true;
            }

            if (!XLSheetPoint.TryParse(input[..separatorIndex], out var first) ||
                !XLSheetPoint.TryParse(input[(separatorIndex + 1)..], out var second) ||
                first.Column > second.Column || first.Row > second.Row)
            {
                area = default;
                return false;
            }

            area = new XLSheetRange(first, second);
            return true;
        }

        /// <summary>
        /// Write the sheet range to the span. If range has only one cell, write only the cell.
        /// </summary>
        /// <param name="output">Must be at least 21 chars long.</param>
        /// <returns>Number of written characters.</returns>
        public int Format(Span<char> output)
        {
            if (FirstPoint == LastPoint)
                return FirstPoint.Format(output);

            var firstPointLen = FirstPoint.Format(output);
            output[firstPointLen] = ':';
            var lastPointLen = LastPoint.Format(output.Slice(firstPointLen + 1));
            return firstPointLen + 1 + lastPointLen;
        }

        public override String ToString()
        {
            Span<char> text = stackalloc char[21];
            var len = Format(text);
            return text.Slice(0, len).ToString();
        }

        /// <summary>
        /// Return a range that contains all cells below the current range.
        /// </summary>
        /// <exception cref="InvalidOperationException">The range touches the bottom border of the sheet.</exception>
        internal XLSheetRange BelowRange()
        {
            return BelowRange(XLHelper.MaxRowNumber);
        }

        /// <summary>
        /// Get a range below the current one <paramref name="rows"/> rows.
        /// If there isn't enough rows, use as many as possible.
        /// </summary>
        /// <exception cref="InvalidOperationException">The range touches the bottom border of the sheet.</exception>
        internal XLSheetRange BelowRange(int rows)
        {
            if (LastPoint.Row >= XLHelper.MaxRowNumber)
                throw new InvalidOperationException("No cells below.");

            rows = Math.Min(rows, XLHelper.MaxRowNumber - LastPoint.Row);
            return new XLSheetRange(
                new XLSheetPoint(LastPoint.Row + 1, FirstPoint.Column),
                new XLSheetPoint(LastPoint.Row + rows, LastPoint.Column));
        }

        /// <summary>
        /// Return a range that contains all cells to the right of the range.
        /// </summary>
        /// <exception cref="InvalidOperationException">The range touches the right border of the sheet.</exception>
        internal XLSheetRange RightRange()
        {
            if (LastPoint.Column == XLHelper.MaxColumnNumber)
                throw new InvalidOperationException("No cells to the left.");

            return new XLSheetRange(
                new XLSheetPoint(FirstPoint.Row, LastPoint.Column + 1),
                new XLSheetPoint(LastPoint.Row, XLHelper.MaxColumnNumber));
        }

        /// <summary>
        /// Return a range that contains additional number of rows below.
        /// </summary>
        internal XLSheetRange ExtendBelow(int rows)
        {
            Debug.Assert(rows >= 0);
            var row = Math.Min(LastPoint.Row + rows, XLHelper.MaxRowNumber);
            return new XLSheetRange(FirstPoint, new XLSheetPoint(row, LastPoint.Column));
        }

        /// <summary>
        /// Return a range that contains additional number of columns to the right.
        /// </summary>
        internal XLSheetRange ExtendRight(int columns)
        {
            Debug.Assert(columns >= 0);
            var column = Math.Min(LastPoint.Column + columns, XLHelper.MaxColumnNumber);
            return new XLSheetRange(FirstPoint, new XLSheetPoint(LastPoint.Row, column));
        }

        internal static XLSheetRange FromRangeAddress<T>(T address)
            where T : IXLRangeAddress
        {
            var firstPoint = XLSheetPoint.FromAddress(address.FirstAddress);
            var lastPoint = XLSheetPoint.FromAddress(address.LastAddress);
            if (firstPoint.Row > lastPoint.Row || firstPoint.Column > lastPoint.Column)
                return new XLSheetRange(lastPoint, firstPoint);

            return new XLSheetRange(firstPoint, lastPoint);
        }

        public bool Contains(XLSheetPoint point)
        {
            return
                point.Row >= FirstPoint.Row && point.Row <= LastPoint.Row &&
                point.Column >= FirstPoint.Column && point.Column <= LastPoint.Column;
        }

        /// <summary>
        /// Create a new range from this one by taking a number of rows from the bottom row up.
        /// </summary>
        /// <param name="rows">How many rows to take, must be at least one.</param>
        public XLSheetRange SliceFromBottom(int rows)
        {
            if (rows < 1)
                throw new ArgumentOutOfRangeException();

            return new XLSheetRange(new XLSheetPoint(BottomRow - rows + 1, FirstPoint.Column), LastPoint);
        }

        /// <summary>
        /// Create a new range from this one by taking a number of rows from the top row down.
        /// </summary>
        /// <param name="rows">How many rows to take, must be at least one.</param>
        public XLSheetRange SliceFromTop(int rows)
        {
            if (rows < 1)
                throw new ArgumentOutOfRangeException();

            return new XLSheetRange(FirstPoint, new XLSheetPoint(TopRow + rows - 1, LastPoint.Column));
        }

        /// <summary>
        /// Create a new range from this one by taking a number of rows from the left column to the right.
        /// </summary>
        /// <param name="columns">How many columns to take, must be at least one.</param>
        public XLSheetRange SliceFromLeft(int columns)
        {
            if (columns < 1)
                throw new ArgumentOutOfRangeException();

            return new XLSheetRange(FirstPoint, new XLSheetPoint(LastPoint.Row, LeftColumn + columns - 1));
        }

        /// <summary>
        /// Create a new range from this one by taking a number of rows from the bottom row up.
        /// </summary>
        /// <param name="columns">How many columns to take, must be at least one.</param>
        public XLSheetRange SliceFromRight(int columns)
        {
            if (columns < 1)
                throw new ArgumentOutOfRangeException();

            return new XLSheetRange(new XLSheetPoint(FirstPoint.Row, RightColumn - columns + 1), LastPoint);
        }

        /// <summary>
        /// Create a new sheet range that is a result of range operator (<c>:</c>)
        /// of this sheet range and <paramref name="otherRange"/>
        /// </summary>
        /// <param name="otherRange">The other range.</param>
        /// <returns>A range that contains both this range and <paramref name="otherRange"/>.</returns>
        public XLSheetRange Range(XLSheetRange otherRange)
        {
            var topRow = Math.Min(TopRow, otherRange.TopRow);
            var leftColumn = Math.Min(LeftColumn, otherRange.LeftColumn);
            var bottomRow = Math.Max(BottomRow, otherRange.BottomRow);
            var rightColumn = Math.Max(RightColumn, otherRange.RightColumn);
            return new XLSheetRange(topRow, leftColumn, bottomRow, rightColumn);
        }

        /// <summary>
        /// Does this range intersects with <paramref name="other"/>.
        /// </summary>
        /// <returns><c>true</c> if intersects, <c>false</c> otherwise.</returns>
        internal bool Intersects(XLSheetRange other)
        {
            return Intersect(other) is not null;
        }

        /// <summary>
        /// Do an intersection between this range and other range.
        /// </summary>
        /// <param name="other">Other range.</param>
        /// <returns>The intersection range if it exists and is non-empty or null, if intersection doesn't exist.</returns>
        internal XLSheetRange? Intersect(XLSheetRange other)
        {
            var leftColumn = Math.Max(LeftColumn, other.LeftColumn);
            var rightColumn = Math.Min(RightColumn, other.RightColumn);
            var topRow = Math.Max(TopRow, other.TopRow);
            var bottomRow = Math.Min(BottomRow, other.BottomRow);

            if (bottomRow < topRow || rightColumn < leftColumn)
                return null;

            return new XLSheetRange(topRow, leftColumn, bottomRow, rightColumn);
        }

        /// <summary>
        /// Does this range overlaps the <paramref name="otherRange"/>?
        /// </summary>
        internal bool Overlaps(XLSheetRange otherRange)
        {
            return TopRow <= otherRange.TopRow &&
                RightColumn >= otherRange.RightColumn &&
                BottomRow >= otherRange.BottomRow &&
                LeftColumn <= otherRange.LeftColumn;
        }

        /// <summary>
        /// Does range cover all rows, from top row to bottom row of a sheet.
        /// </summary>
        internal bool IsEntireColumn()
        {
            return TopRow == 1 && BottomRow == XLHelper.MaxRowNumber;
        }

        /// <summary>
        /// Does range cover all columns, from first to last column of a sheet.
        /// </summary>
        public bool IsEntireRow()
        {
            return LeftColumn == 1 && RightColumn == XLHelper.MaxColumnNumber;
        }

        /// <summary>
        /// Return a new range that has the same size as the current one,
        /// </summary>
        /// <param name="topLeftCorner">New top left coordinate of returned range.</param>
        /// <returns>New range.</returns>
        internal XLSheetRange At(XLSheetPoint topLeftCorner)
        {
            var bottomRightCorner = topLeftCorner.ShiftColumn(Width - 1).ShiftRow(Height - 1);
            return new XLSheetRange(topLeftCorner, bottomRightCorner);
        }

        /// <summary>
        /// Return a new range that has been shifted in vertical direction by <paramref name="rowShift"/>.
        /// </summary>
        /// <param name="rowShift">By how much to shift the range, positive - downwards, negative - upwards.</param>
        /// <returns>Newly created area.</returns>
        internal XLSheetRange ShiftRows(int rowShift)
        {
            var topLeftCorner = FirstPoint.ShiftRow(rowShift);
            var bottomRightCorner = LastPoint.ShiftRow(rowShift);
            return new XLSheetRange(topLeftCorner, bottomRightCorner);
        }

        /// <summary>
        /// Return a new range that has been shifted in vertical direction by <paramref name="rowShift"/>.
        /// If the shifted area is out of sheet bounds, clip part that is out.
        /// </summary>
        /// <param name="rowShift">How many rows to shift.</param>
        /// <returns>Shifted clipped area or <c>null</c> if area was shifted completely out of a sheet.</returns>
        internal XLSheetRange? ShiftRowsAndClip(int rowShift)
        {
            var shiftedTop = TopRow + rowShift;
            if (shiftedTop > XLHelper.MaxRowNumber)
                return null;

            var shiftedBottom = BottomRow + rowShift;
            if (shiftedBottom < XLHelper.MinRowNumber)
                return null;

            var clippedTop = Math.Max(shiftedTop, XLHelper.MinRowNumber);
            var clippedBottom = Math.Min(shiftedBottom, XLHelper.MaxRowNumber);

            return new XLSheetRange(clippedTop, LeftColumn, clippedBottom, RightColumn);
        }

        /// <summary>
        /// Return a new range that has been shifted in horizontal direction by <paramref name="columnShift"/>.
        /// </summary>
        /// <param name="columnShift">By how much to shift the range, positive - rightward, negative - leftward.</param>
        /// <returns>Newly created area.</returns>
        internal XLSheetRange ShiftColumns(int columnShift)
        {
            var topLeftCorner = FirstPoint.ShiftColumn(columnShift);
            var bottomRightCorner = LastPoint.ShiftColumn(columnShift);
            return new XLSheetRange(topLeftCorner, bottomRightCorner);
        }

        /// <summary>
        /// Return a new range that has been shifted in horizontal direction by <paramref name="columnShift"/>.
        /// If the shifted area is out of sheet bounds, clip part that is out.
        /// </summary>
        /// <param name="columnShift">How many columns to shift.</param>
        /// <returns>Shifted clipped area or <c>null</c> if area was shifted completely out of a sheet.</returns>
        internal XLSheetRange? ShiftColumnsAndClip(int columnShift)
        {
            var shiftedLeft = LeftColumn + columnShift;
            if (shiftedLeft > XLHelper.MaxColumnNumber)
                return null;

            var shiftedRight = RightColumn + columnShift;
            if (shiftedRight < XLHelper.MinColumnNumber)
                return null;

            var clippedLeft = Math.Max(shiftedLeft, XLHelper.MinColumnNumber);
            var clippedRight = Math.Min(shiftedRight, XLHelper.MaxColumnNumber);

            return new XLSheetRange(TopRow, clippedLeft, BottomRow, clippedRight);
        }

        public IEnumerator<XLSheetPoint> GetEnumerator()
        {
            for (var row = TopRow; row <= BottomRow; ++row)
            {
                for (var col = LeftColumn; col <= RightColumn; ++col)
                {
                    yield return new XLSheetPoint(row, col);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Calculate size and position of the area when another area is inserted into a sheet.
        /// </summary>
        /// <param name="insertedArea">Inserted area.</param>
        /// <param name="result">The result, might be <c>null</c> as a valid result if area is pushed out.</param>
        /// <returns><c>true</c> if results wasn't partially shifted.</returns>
        internal bool TryInsertAreaAndShiftRight(XLSheetRange insertedArea, out XLSheetRange? result)
        {
            // Inserted fully upward, downward or to the right
            if (insertedArea.BottomRow < TopRow ||
                insertedArea.TopRow > BottomRow ||
                insertedArea.LeftColumn > RightColumn)
            {
                result = this;
                return true;
            }

            var fullyOverlaps = insertedArea.TopRow <= TopRow &&
                                insertedArea.BottomRow >= BottomRow;
            if (!fullyOverlaps)
            {
                result = null;
                return false;
            }

            // Are is effectively inserted into a seam at the left column of the insertedArea
            if (insertedArea.LeftColumn <= LeftColumn)
            {
                // Area is completely pushed out
                if (LeftColumn + insertedArea.Width > XLHelper.MaxColumnNumber)
                {
                    result = null;
                    return true;
                }

                // Area is partially pushed out
                if (RightColumn + insertedArea.Width > XLHelper.MaxColumnNumber)
                {
                    var pushedOutColsCount = RightColumn + insertedArea.Width - XLHelper.MaxColumnNumber;
                    var keepCols = Width - pushedOutColsCount;
                    var resized = SliceFromLeft(keepCols);
                    result = resized.ShiftColumns(insertedArea.Width);
                    return true;
                }

                // Not pushed out = only shift
                result = ShiftColumns(insertedArea.Width);
                return true;
            }

            result = ExtendRight(insertedArea.Width);
            return true;
        }

        /// <summary>
        /// Calculate size and position of the area when another area is inserted into a sheet.
        /// </summary>
        /// <param name="insertedArea">Inserted area.</param>
        /// <param name="result">The result, might be <c>null</c> as a valid result if area is pushed out.</param>
        /// <returns><c>true</c> if results wasn't partially shifted.</returns>
        internal bool TryInsertAreaAndShiftDown(XLSheetRange insertedArea, out XLSheetRange? result)
        {
            // Inserted fully to the left, to the right or below
            if (insertedArea.RightColumn < LeftColumn ||
                insertedArea.LeftColumn > RightColumn ||
                insertedArea.TopRow > BottomRow)
            {
                result = this;
                return true;
            }

            var fullyOverlaps = insertedArea.LeftColumn <= LeftColumn &&
                                insertedArea.RightColumn >= RightColumn;
            if (!fullyOverlaps)
            {
                result = null;
                return false;
            }

            // Are is effectively inserted into a seam at the top row of the insertedArea
            if (insertedArea.TopRow <= TopRow)
            {
                // Area is completely pushed out
                if (TopRow + insertedArea.Height > XLHelper.MaxRowNumber)
                {
                    result = null;
                    return true;
                }

                // Area is partially pushed out
                if (BottomRow + insertedArea.Height > XLHelper.MaxRowNumber)
                {
                    var pushedOutRowsCount = BottomRow + insertedArea.Height - XLHelper.MaxRowNumber;
                    var keepRows = Height - pushedOutRowsCount;
                    var resized = SliceFromTop(keepRows);
                    result = resized.ShiftRows(insertedArea.Height);
                    return true;
                }

                // Not pushed out = only shift
                result = ShiftRows(insertedArea.Height);
                return true;
            }

            result = ExtendBelow(insertedArea.Height);
            return true;
        }

        /// <summary>
        /// Take the area and reposition it as if the <paramref name="deletedArea"/> was removed
        /// from sheet. If cells the left of the area are deleted, the area shifts to the left.
        /// If <paramref name="deletedArea"/> is within the area, the width of the area decreases.
        /// </summary>
        /// <remarks>
        /// If the method returns <c>false</c>, there is a partial cover and it's up to you to
        /// decide what to do.
        /// </remarks>
        /// <returns>
        /// The <paramref name="result"/> has a value <c>null</c> if the range was completely
        /// removed by <paramref name="deletedArea"/>.
        /// </returns>
        internal bool TryDeleteAreaAndShiftLeft(XLSheetRange deletedArea, out XLSheetRange? result)
        {
            // Deleted area is fully upwards, downwards or to the right of this area.
            if (deletedArea.BottomRow < TopRow ||
                deletedArea.TopRow > BottomRow ||
                deletedArea.LeftColumn > RightColumn)
            {
                result = this;
                return true;
            }

            var coversWidth = deletedArea.LeftColumn <= LeftColumn &&
                              deletedArea.RightColumn >= RightColumn;
            var coversHeight = deletedArea.TopRow <= TopRow &&
                               deletedArea.BottomRow >= BottomRow;
            var fullyCovered = coversWidth && coversHeight;
            if (fullyCovered)
            {
                result = null;
                return true;
            }

            // When a slice form a top/bottom is deleted, the rest doesn't move.
            // There is no split either. Whole slice is just removed.
            var deletedTopSlice = coversWidth &&
                                  deletedArea.TopRow <= TopRow &&
                                  deletedArea.BottomRow < BottomRow;
            if (deletedTopSlice)
            {
                var sliceRows = BottomRow - deletedArea.BottomRow;
                result = SliceFromBottom(sliceRows);
                return true;
            }

            var deletedBottomSlice = coversWidth &&
                                     deletedArea.BottomRow >= BottomRow &&
                                     deletedArea.TopRow > TopRow;
            if (deletedBottomSlice)
            {
                var sliceRows = deletedArea.TopRow - TopRow;
                result = SliceFromTop(sliceRows);
                return true;
            }

            // Slice cases were already dealt with, anything that doesn't cover height would cause split
            if (!coversHeight)
            {
                result = null;
                return false;
            }

            var deletesColumnsToLeft = deletedArea.LeftColumn < LeftColumn;
            var deletesColumnsOfArea = deletedArea.LeftColumn <= RightColumn &&
                                       deletedArea.RightColumn >= LeftColumn;
            var repositioned = this;
            if (deletesColumnsOfArea)
            {
                // Decrease width of repositioned area
                var left = Math.Max(deletedArea.LeftColumn, repositioned.LeftColumn);
                var right = Math.Min(deletedArea.RightColumn, repositioned.RightColumn);

                var columnsToDelete = right - left + 1;
                var newWidth = repositioned.Width - columnsToDelete;
                if (newWidth == 0)
                {
                    result = null;
                    return true;
                }

                repositioned = repositioned.SliceFromLeft(newWidth);
            }

            if (deletesColumnsToLeft)
            {
                // There are some deleted columns to the left of the area -> shift left
                var deletedLastColumnsOutwards = Math.Min(repositioned.LeftColumn - 1, deletedArea.RightColumn);

                var shiftLeft = deletedLastColumnsOutwards - deletedArea.LeftColumn + 1;
                repositioned = repositioned.ShiftColumns(-shiftLeft);
            }

            result = repositioned;
            return true;
        }

        /// <summary>
        /// Take the area and reposition it as if the <paramref name="deletedArea"/> was removed
        /// from sheet. If cells upward of the area are deleted, the area shifts to the upward.
        /// If <paramref name="deletedArea"/> is within the area, the height of the area decreases.
        /// </summary>
        /// <remarks>
        /// If the method returns <c>false</c>, there is a partial cover and it's up to you to
        /// decide what to do.
        /// </remarks>
        /// <returns>
        /// The <paramref name="result"/> has a value <c>null</c> if the range was completely
        /// removed by <paramref name="deletedArea"/>.
        /// </returns>
        internal bool TryDeleteAreaAndShiftUp(XLSheetRange deletedArea, out XLSheetRange? result)
        {
            // Deleted area is fully on left, right or bottom side of this area.
            if (deletedArea.RightColumn < LeftColumn ||
                deletedArea.LeftColumn > RightColumn ||
                deletedArea.TopRow > BottomRow)
            {
                result = this;
                return true;
            }

            var coversWidth = deletedArea.LeftColumn <= LeftColumn &&
                              deletedArea.RightColumn >= RightColumn;
            var coversHeight = deletedArea.TopRow <= TopRow &&
                               deletedArea.BottomRow >= BottomRow;
            var fullyCovered = coversWidth && coversHeight;
            if (fullyCovered)
            {
                result = null;
                return true;
            }

            // When a slice form a left/right is deleted, the rest doesn't move.
            // There is no split either. Whole slice is just removed.
            var deletedLeftSlice = coversHeight &&
                                   deletedArea.LeftColumn <= LeftColumn &&
                                   deletedArea.RightColumn < RightColumn;
            if (deletedLeftSlice)
            {
                var sliceColumns = RightColumn - deletedArea.RightColumn;
                result = SliceFromRight(sliceColumns);
                return true;
            }

            var deletedRightSlice = coversHeight &&
                                    deletedArea.RightColumn >= RightColumn &&
                                    deletedArea.LeftColumn > LeftColumn;
            if (deletedRightSlice)
            {
                var sliceRows = deletedArea.LeftColumn - LeftColumn;
                result = SliceFromLeft(sliceRows);
                return true;
            }

            // Slice cases were already dealt with, anything that doesn't cover height would cause split
            var doesntOverlapWidth = deletedArea.LeftColumn > LeftColumn ||
                                     deletedArea.RightColumn < RightColumn;
            if (doesntOverlapWidth)
            {
                result = null;
                return false;
            }

            var deletesRowsAboveArea = deletedArea.TopRow < TopRow;
            var deletesRowsOfArea = deletedArea.TopRow <= BottomRow &&
                                    deletedArea.BottomRow >= TopRow;
            var repositioned = this;
            if (deletesRowsOfArea)
            {
                // Decrease height of repositioned area
                var top = Math.Max(deletedArea.TopRow, repositioned.TopRow);
                var bottom = Math.Min(deletedArea.BottomRow, repositioned.BottomRow);

                var rowsToDelete = bottom - top + 1;
                var newHeight = repositioned.Height - rowsToDelete;
                if (newHeight == 0)
                {
                    result = null;
                    return true;
                }

                repositioned = repositioned.SliceFromTop(newHeight);
            }

            if (deletesRowsAboveArea)
            {
                // There are some deleted rows above the area -> shift up
                var deletedLastRowAboveArea = Math.Min(repositioned.TopRow - 1, deletedArea.BottomRow);

                var shiftUp = deletedLastRowAboveArea - deletedArea.TopRow + 1;
                repositioned = repositioned.ShiftRows(-shiftUp);
            }

            result = repositioned;
            return true;
        }

        /// <summary>
        /// Determine a areas that contain all cells of this area without <paramref name="range"/>
        /// and add them to the <paramref name="nonExcludedAreas"/>.
        /// </summary>
        /// <param name="range">Range to exclude from this one.</param>
        /// <param name="nonExcludedAreas">A list to which add remaining (non-excluded) areas.</param>
        /// <returns>If an area was excluded, the excluded area.</returns>
        internal XLSheetRange? Exclude(XLSheetRange range, List<XLSheetRange> nonExcludedAreas)
        {
            if (Intersect(range) is not { } intersection)
            {
                nonExcludedAreas.Add(this);
                return null;
            }

            // left
            if (LeftColumn < intersection.LeftColumn)
                nonExcludedAreas.Add(new XLSheetRange(TopRow, LeftColumn, BottomRow, intersection.LeftColumn - 1));

            // top
            if (TopRow < intersection.TopRow)
                nonExcludedAreas.Add(new XLSheetRange(TopRow, intersection.LeftColumn, intersection.TopRow - 1, intersection.RightColumn));

            // bottom
            if (BottomRow > intersection.BottomRow)
                nonExcludedAreas.Add(new XLSheetRange(intersection.BottomRow + 1, intersection.LeftColumn, BottomRow, intersection.RightColumn));

            // right
            if (RightColumn > intersection.RightColumn)
                nonExcludedAreas.Add(new XLSheetRange(TopRow, intersection.RightColumn + 1, BottomRow, RightColumn));

            return intersection;
        }

        /// <summary>
        /// Return an area that has dimensions as if columns were inserted at <paramref name="insertedLeftColumn"/>.
        /// Mimics Excel behavior.
        /// </summary>
        /// <param name="insertedLeftColumn">A position where columns are inserted.</param>
        /// <param name="insertedWidth">How many columns were inserted.</param>
        internal XLSheetRange? ShiftOrExtendRight(int insertedLeftColumn, int insertedWidth)
        {
            Debug.Assert(insertedWidth >= 0);

            // Area inserted at the right edge extends - that is the reason for - 1
            if (RightColumn < insertedLeftColumn - 1)
            {
                // inserted is to the right of area -> no shift
                return this;
            }

            if (LeftColumn >= insertedLeftColumn)
            {
                // Inserted is to the left of affected area -> shift
                return ShiftColumnsAndClip(insertedWidth);
            }

            // inserted is in the middle of affected: affectedLeft < insertedLeft <= affectedRight
            return ExtendRight(insertedWidth);
        }

        /// <summary>
        /// Return an area that has dimensions as if a rows were inserted at <paramref name="insertedTopRow"/>.
        /// Mimics Excel behavior.
        /// </summary>
        /// <param name="insertedTopRow">A position where rows are inserted.</param>
        /// <param name="insertedHeight">How many rows were inserted.</param>
        internal XLSheetRange? ShiftOrExtendDown(int insertedTopRow, int insertedHeight)
        {
            Debug.Assert(insertedHeight >= 0);

            // Area inserted at the bottom edge extends - that is the reason for - 1
            if (BottomRow < insertedTopRow - 1)
            {
                // inserted is below the area -> no shift
                return this;
            }

            if (TopRow >= insertedTopRow)
            {
                // Inserted is above the area -> shift
                return ShiftRowsAndClip(insertedHeight);
            }

            // inserted is in the middle of affected: affectedTop < insertedTop <= affectedBottom
            return ExtendBelow(insertedHeight);
        }

        /// <summary>
        /// Return an area that has dimensions as if a rows were deleted from <paramref name="deletedTopRow"/>.
        /// Mimics Excel behavior.
        /// </summary>
        /// <param name="deletedTopRow">A position from which where rows are deleted.</param>
        /// <param name="deletedHeight">How many rows were deleted.</param>
        internal XLSheetRange? ShiftOrShrinkUp(int deletedTopRow, int deletedHeight)
        {
            Debug.Assert(deletedHeight >= 0);
            if (BottomRow < deletedTopRow || deletedHeight == 0)
            {
                // deleted is below the area -> no shift or shrink
                return this;
            }

            var deletedBottomRow = deletedTopRow + deletedHeight - 1;
            if (deletedBottomRow < TopRow)
            {
                // Deleted area is completely above the area -> only shift
                return ShiftRows(-deletedHeight);
            }

            // Shrink by how much deletedArea and area overlap
            var shrink = Math.Min(BottomRow, deletedBottomRow) - Math.Max(TopRow, deletedTopRow) + 1;
            if (shrink == Height)
                return null;

            var shift = Math.Max(TopRow - deletedTopRow, 0);
            var shifted = ShiftRows(-shift);
            return new XLSheetRange(shifted.TopRow, shifted.LeftColumn, shifted.BottomRow - shrink, shifted.RightColumn);
        }

        /// <summary>
        /// Return an area that has dimensions as if a column were deleted from <paramref name="deletedLeftColumn"/>.
        /// Mimics Excel behavior.
        /// </summary>
        /// <param name="deletedLeftColumn">A position from which where columns are deleted.</param>
        /// <param name="deletedWidth">How many columns were deleted.</param>
        internal XLSheetRange? ShiftOrShrinkLeft(int deletedLeftColumn, int deletedWidth)
        {
            Debug.Assert(deletedWidth >= 0);
            if (RightColumn < deletedLeftColumn || deletedWidth == 0)
            {
                // deleted is to the right of area -> no shift or shrink
                return this;
            }

            var deletedRightColumn = deletedLeftColumn + deletedWidth - 1;
            if (deletedRightColumn < LeftColumn)
            {
                // Deleted area is completely to left of area -> only shift
                return ShiftColumns(-deletedWidth);
            }

            // Shrink by how much deletedArea and area overlap
            var shrink = Math.Min(RightColumn, deletedRightColumn) - Math.Max(LeftColumn, deletedLeftColumn) + 1;
            if (shrink == Width)
                return null;

            var shift = Math.Max(LeftColumn - deletedLeftColumn, 0);
            var shifted = ShiftColumns(-shift);
            return new XLSheetRange(shifted.TopRow, shifted.LeftColumn, shifted.BottomRow, shifted.RightColumn - shrink);
        }
    }
}
