using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A specification of an area (rectangular range) of a sheet.
    /// </summary>
    internal readonly struct SheetArea : IEquatable<SheetArea>, IEnumerable<SheetPoint>
    {
        /// <summary>
        /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// An area in the sheet.
        /// </summary>
        public readonly Area Area;

        public SheetArea(String name, Area area)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException(nameof(name));

            Name = name;
            Area = area;
        }

        public static bool operator ==(SheetArea lhs, SheetArea rhs) => lhs.Equals(rhs);

        public static bool operator !=(SheetArea lhs, SheetArea rhs) => !(lhs == rhs);

        public IEnumerator<SheetPoint> GetEnumerator()
        {
            for (var row = Area.TopRow; row <= Area.BottomRow; ++row)
            {
                for (var col = Area.LeftColumn; col <= Area.RightColumn; ++col)
                {
                    yield return new SheetPoint(Name, row, col);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal static SheetArea From(IXLRange range)
        {
            if (range.Worksheet is null)
                throw new ArgumentException("Range doesn't contain sheet.", nameof(range));

            return new SheetArea(range.Worksheet.Name, Area.FromRangeAddress(range.RangeAddress));
        }

        internal static SheetArea From(XLRangeAddress address)
        {
            if (address.Worksheet is null)
                throw new ArgumentException("Range doesn't contain sheet.", nameof(address));

            return new SheetArea(address.Worksheet.Name, Area.FromRangeAddress(address));
        }

        public bool Equals(SheetArea other)
        {
            return Area == other.Area && XLHelper.SheetComparer.Equals(Name, other.Name);
        }

        public override bool Equals(object? obj)
        {
            return obj is SheetArea other && Equals(other);
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
        public SheetArea? Intersect(SheetArea other)
        {
            if (!XLHelper.SheetComparer.Equals(Name, other.Name))
                return null;

            var intersectionRange = Area.Intersect(other.Area);
            if (intersectionRange is null)
                return null;

            return new SheetArea(Name, intersectionRange.Value);
        }

        public void Deconstruct(out string sheetName, out Area area)
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
