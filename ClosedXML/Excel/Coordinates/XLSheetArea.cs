﻿using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A specification of an area (rectangular range) of a sheet.
    /// </summary>
    internal readonly struct XLSheetArea : IEquatable<XLSheetArea>
    {
        /// <summary>
        /// Name of the sheet. Sheet may exist or not (e.g. deleted).
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// An area in the sheet.
        /// </summary>
        public readonly XLSheetRange Area;

        public XLSheetArea(String name, XLSheetRange area)
        {
            Name = name;
            Area = area;
        }

        public bool Equals(XLSheetArea other)
        {
            return Area == other.Area && XLHelper.SheetComparer.Equals(Name, other.Name);
        }

        public override bool Equals(object? obj)
        {
            return obj is XLSheetArea other && Equals(other);
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
        public XLSheetArea? Intersect(XLSheetArea other)
        {
            if (!XLHelper.SheetComparer.Equals(Name, other.Name))
                return null;

            var intersectionRange = Area.Intersect(other.Area);
            if (intersectionRange is null)
                return null;

            return new XLSheetArea(Name, intersectionRange.Value);
        }
    }
}
