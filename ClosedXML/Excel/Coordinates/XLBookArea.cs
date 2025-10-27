using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A specification of an area (rectangular range) of a sheet.
    /// </summary>
    internal readonly struct XLBookArea : IEquatable<XLBookArea>, IEnumerable<XLBookPoint>
    {
        /// <summary>
        /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// An area in the sheet.
        /// </summary>
        public readonly XLSheetRange Area;

        public XLBookArea(String name, XLSheetRange area)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException(nameof(name));

            Name = name;
            Area = area;
        }

        public static bool operator ==(XLBookArea lhs, XLBookArea rhs) => lhs.Equals(rhs);

        public static bool operator !=(XLBookArea lhs, XLBookArea rhs) => !(lhs == rhs);

        public IEnumerator<XLBookPoint> GetEnumerator()
        {
            for (var row = Area.TopRow; row <= Area.BottomRow; ++row)
            {
                for (var col = Area.LeftColumn; col <= Area.RightColumn; ++col)
                {
                    yield return new XLBookPoint(Name, row, col);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal static XLBookArea From(IXLRange range)
        {
            if (range.Worksheet is null)
                throw new ArgumentException("Range doesn't contain sheet.", nameof(range));

            return new XLBookArea(range.Worksheet.Name, XLSheetRange.FromRangeAddress(range.RangeAddress));
        }

        internal static XLBookArea From(XLRangeAddress address)
        {
            if (address.Worksheet is null)
                throw new ArgumentException("Range doesn't contain sheet.", nameof(address));

            return new XLBookArea(address.Worksheet.Name, XLSheetRange.FromRangeAddress(address));
        }

        public bool Equals(XLBookArea other)
        {
            return Area == other.Area && XLHelper.SheetComparer.Equals(Name, other.Name);
        }

        public override bool Equals(object? obj)
        {
            return obj is XLBookArea other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (XLHelper.SheetComparer.GetHashCode(Name) * 397) ^ Area.GetHashCode();
            }
        }

        /// <summary>
        /// Perform an intersection.
        /// </summary>
        /// <param name="other">The area that is being intersected with this one.</param>
        /// <returns>The intersection (=same sheet and has non-empty intersection) or null if intersection isn't possible.</returns>
        public XLBookArea? Intersect(XLBookArea other)
        {
            if (!XLHelper.SheetComparer.Equals(Name, other.Name))
                return null;

            var intersectionRange = Area.Intersect(other.Area);
            if (intersectionRange is null)
                return null;

            return new XLBookArea(Name, intersectionRange.Value);
        }

        public void Deconstruct(out string sheetName, out XLSheetRange area)
        {
            sheetName = Name;
            area = Area;
        }

        public override string ToString()
        {
            return $"{Name}!{Area}";
        }
    }
}
