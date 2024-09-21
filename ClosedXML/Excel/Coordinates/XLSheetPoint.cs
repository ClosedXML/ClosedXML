using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An point (address) in a worksheet, an equivalent of <c>ST_CellRef</c>.
    /// </summary>
    /// <remarks>Unlike the XLAddress, sheet can never be invalid.</remarks>
    [DebuggerDisplay("{XLHelper.GetColumnLetterFromNumber(Column)+Row}")]
    internal readonly struct XLSheetPoint : IEquatable<XLSheetPoint>, IComparable<XLSheetPoint>
    {
        public XLSheetPoint(Int32 row, Int32 column)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 1-based row number in a sheet.
        /// </summary>
        public readonly Int32 Row;

        /// <summary>
        /// 1-based column number in a sheet.
        /// </summary>
        public readonly Int32 Column;

        public static implicit operator XLSheetRange(XLSheetPoint point)
        {
            return new XLSheetRange(point);
        }

        public override bool Equals(object obj)
        {
            return obj is XLSheetPoint point && Equals(point);
        }

        public bool Equals(XLSheetPoint other)
        {
            return Row == other.Row && Column == other.Column;
        }

        public override int GetHashCode()
        {
            return (Row * -1) ^ Column;
        }

        public static bool operator ==(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row == b.Row && a.Column == b.Column;
        }

        public static bool operator !=(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row != b.Row || a.Column != b.Column;
        }

        /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
        public static XLSheetPoint Parse(String text) => Parse(text.AsSpan());

        /// <summary>
        /// Parse point per type <c>ST_CellRef</c> from
        /// <a href="https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db11a912-b1cb-4dff-b46d-9bedfd10cef0">2.1.1108 Part 4 Section 3.18.8, ST_CellRef (Cell Reference)</a>
        /// </summary>
        /// <param name="input">Input text</param>
        /// <exception cref="FormatException">If the input doesn't match expected grammar.</exception>
        public static XLSheetPoint Parse(ReadOnlySpan<char> input)
        {
            // Don't reuse inefficient logic from XLAddress
            if (input.Length < 2)
                throw new FormatException($"Length is less than two ('{input.ToString()}').");

            var i = 0;
            var c = input[i++];
            if (!IsLetter(c))
                throw new FormatException($"Doesn't start with a letter ('{input.ToString()}').");

            var columnIndex = c - 'A' + 1;
            while (i < input.Length && IsLetter(c = input[i]))
            {
                columnIndex = columnIndex * 26 + c - 'A' + 1;
                i++;
            }

            if (i > 3)
                throw new FormatException($"Input contains more than three letters ('{input.ToString()}').");

            if (i == input.Length)
                throw new FormatException($"Input doesn't contain row number ('{input.ToString()}').");

            // Everything else must be digits
            c = input[i++];

            // First letter can't be 0
            if (c is < '1' or > '9')
                throw new FormatException($"Row must start with a non-zero digit ('{input.ToString()}').");

            var rowIndex = c - '0';
            while (i < input.Length && IsDigit(c = input[i]))
            {
                rowIndex = rowIndex * 10 + c - '0';
                i++;
            }

            if (i != input.Length)
                throw new FormatException($"Input contains unexpected characters ('{input.ToString()}').");

            if (rowIndex > XLHelper.MaxRowNumber || columnIndex > XLHelper.MaxColumnNumber)
                throw new FormatException($"Address out of bounds ('{input.ToString()}').");

            return new XLSheetPoint(rowIndex, columnIndex);

            static bool IsLetter(char c) => c is >= 'A' and <= 'Z';
            static bool IsDigit(char c) => c is >= '0' and <= '9';
        }

        /// <summary>
        /// Write the sheet point as a reference to the span (e.g. <c>A1</c>).
        /// </summary>
        /// <param name="output">Must be at least 10 chars long</param>
        /// <returns>Number of chars </returns>
        public int Format(Span<char> output)
        {
            var columnLetters = XLHelper.GetColumnLetterFromNumber(Column);
            for (var i = 0; i < columnLetters.Length; ++i)
                output[i] = columnLetters[i];

            var digitCount = GetDigitCount(Row);
            var rowRemainder = Row;
            var formattedLength = digitCount + columnLetters.Length;
            for (var i = formattedLength - 1; i >= columnLetters.Length; --i)
            {
                var digit = rowRemainder % 10;
                rowRemainder /= 10;
                output[i] = (char)(digit + '0');
            }

            return formattedLength;
        }

        public override String ToString()
        {
            Span<char> text = stackalloc char[10];
            var len = Format(text);
            return text.Slice(0, len).ToString();
        }

        private static int GetDigitCount(int n)
        {
            if (n < 10L) return 1;
            if (n < 100L) return 2;
            if (n < 1000L) return 3;
            if (n < 10000L) return 4;
            if (n < 100000L) return 5;
            if (n < 1000000L) return 6;
            return 7; // Row can't have more digits
        }

        /// <summary>
        /// Create a sheet point from the address. Workbook is ignored.
        /// </summary>
        public static XLSheetPoint FromAddress(IXLAddress address)
            => new(address.RowNumber, address.ColumnNumber);

        public int CompareTo(XLSheetPoint other)
        {
            var rowComparison = Row.CompareTo(other.Row);
            if (rowComparison != 0)
                return rowComparison;

            return Column.CompareTo(other.Column);
        }

        /// <summary>
        /// Is the point within the range or below the range?
        /// </summary>
        internal bool InRangeOrBelow(in XLSheetRange range)
        {
            return Row >= range.FirstPoint.Row &&
                   Column >= range.FirstPoint.Column &&
                   Column <= range.LastPoint.Column;
        }

        /// <summary>
        /// Is the point within the range or to the left of the range?
        /// </summary>
        internal bool InRangeOrToLeft(in XLSheetRange range)
        {
            return Column >= range.FirstPoint.Column &&
                   Row >= range.FirstPoint.Row &&
                   Row <= range.LastPoint.Row;
        }

        /// <summary>
        /// Return a new point that has its row coordinate shifted by <paramref name="rowShift"/>.
        /// </summary>
        /// <param name="rowShift">How many rows will new point be shifted. Positive - new point
        ///     is downwards, negative - new point is upwards relative to the current point.</param>
        /// <returns>Shifted point.</returns>
        internal XLSheetPoint ShiftRow(int rowShift)
        {
            return new XLSheetPoint(Row + rowShift, Column);
        }

        /// <summary>
        /// Return a new point that has its column coordinate shifted by <paramref name="columnShift"/>.
        /// </summary>
        /// <param name="columnShift">How many columns will new point be shifted. Positive - new
        ///     point is to the right, negative - new point is to the left.</param>
        /// <returns>Shifted point.</returns>
        internal XLSheetPoint ShiftColumn(int columnShift)
        {
            return new XLSheetPoint(Row, Column + columnShift);
        }
    }
}
