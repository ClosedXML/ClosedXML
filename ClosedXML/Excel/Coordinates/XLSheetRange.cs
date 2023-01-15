using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A representation of a <c>ST_Ref</c>, i.e. an area in a sheet (no reference to teh sheet).
    /// </summary>
    internal readonly struct XLSheetRange : IEquatable<XLSheetRange>
    {
        public XLSheetRange(XLSheetPoint firstPoint, XLSheetPoint lastPoint)
        {
            FirstPoint = firstPoint;
            LastPoint = lastPoint;
        }

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

        public bool Equals(XLSheetRange other)
        {
            return FirstPoint.Equals(other.FirstPoint) && LastPoint.Equals(other.LastPoint);
        }

        public override int GetHashCode()
        {
            return FirstPoint.GetHashCode() ^ LastPoint.GetHashCode();
        }

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
            var separatorIndex = input.IndexOf(':');
            if (separatorIndex == -1)
            {
                var sheetPoint = XLSheetPoint.Parse(input);
                return new XLSheetRange(sheetPoint, sheetPoint);
            }

            var first = XLSheetPoint.Parse(input.Slice(0, separatorIndex));
            var second = XLSheetPoint.Parse(input.Slice(separatorIndex + 1, input.Length - separatorIndex - 1));
            if (first.Column > second.Column || first.Row > second.Row)
                throw new FormatException($"First reference must have smaller column and row ('{input.ToString()}')");

            return new XLSheetRange(first, second);
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
    }
}
