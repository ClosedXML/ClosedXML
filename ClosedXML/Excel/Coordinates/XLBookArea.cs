using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A specification of an area (rectangular range) of a sheet.
    /// </summary>
    internal readonly struct XLBookArea : IEquatable<XLBookArea>
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

        internal static XLBookArea From(IXLRange range)
        {
            if (range.Worksheet is null)
                throw new ArgumentException("Range doesn't contain sheet.", nameof(range));

            return new XLBookArea(range.Worksheet.Name, XLSheetRange.FromRangeAddress(range.RangeAddress));
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

        public override string ToString()
        {
            return $"{Name}!{Area}";
        }
    }
}
