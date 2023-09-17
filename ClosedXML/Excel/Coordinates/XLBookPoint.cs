using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A single point in a workbook. The book point might point to a deleted
    /// worksheet, so it might be invalid. Make sure it is checked when
    /// determining the properties of the actual data of the point.
    /// </summary>
    internal readonly struct XLBookPoint : IEquatable<XLBookPoint>
    {
        internal XLBookPoint(XLWorksheet sheet, XLSheetPoint point)
            : this(sheet.SheetId, point)
        {
        }

        internal XLBookPoint(uint sheetId, XLSheetPoint point)
        {
            SheetId = sheetId;
            Point = point;
        }

        /// TODO: SheetId doesn't work nicely with renames, but will in the future.
        /// <summary>
        /// A sheet id of a point. Id of a sheet never changes during workbook
        /// lifecycle (<see cref="XLWorksheet.SheetId"/>), but the sheet may be
        /// deleted, making the sheetId and thus book point invalid.
        /// </summary>
        public uint SheetId { get; }

        /// <inheritdoc cref="XLSheetPoint.Row"/>
        public int Row => Point.Column;

        /// <inheritdoc cref="XLSheetPoint.Column"/>
        public int Column => Point.Column;

        /// <summary>
        /// A point in the sheet.
        /// </summary>
        public XLSheetPoint Point { get; }

        public static bool operator ==(XLBookPoint lhs, XLBookPoint rhs) => lhs.Equals(rhs);

        public static bool operator !=(XLBookPoint lhs, XLBookPoint rhs) => !(lhs == rhs);

        public bool Equals(XLBookPoint other)
        {
            return SheetId == other.SheetId && Point.Equals(other.Point);
        }

        public override bool Equals(object? obj)
        {
            return obj is XLBookPoint other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((int)SheetId * 397) ^ Point.GetHashCode();
            }
        }

        public override string ToString()
        {
            return $"[{SheetId}]{Point}";
        }
    }
}
