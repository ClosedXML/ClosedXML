using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An point (address) in a worksheet, an equivalent of <c>ST_CellRef</c>.
    /// </summary>
    /// <remarks>Unlike the XLAddress, sheet can never be invalid.</remarks>
    internal readonly struct XLSheetPoint : IEquatable<XLSheetPoint>
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
    }
}
