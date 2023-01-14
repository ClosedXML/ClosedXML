using System;

namespace ClosedXML.Excel
{
    internal struct XLSheetRange : IEquatable<XLSheetRange>
    {
        public XLSheetRange(XLSheetPoint firstPoint, XLSheetPoint lastPoint)
        {
            FirstPoint = firstPoint;
            LastPoint = lastPoint;
        }

        public readonly XLSheetPoint FirstPoint;
        public readonly XLSheetPoint LastPoint;

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
    }
}
