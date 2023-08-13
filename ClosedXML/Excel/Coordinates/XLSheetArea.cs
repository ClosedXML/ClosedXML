using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A specification of an area (rectangular range) of a sheet.
    /// </summary>
    internal readonly struct XLSheetArea : IEquatable<XLSheetArea>
    {
        /// <summary>
        /// Name of the sheet. Sheet may exist or not.
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
            return Area == other.Area && StringComparer.OrdinalIgnoreCase.Equals(Name, other.Name);
        }

        public override bool Equals(object? obj)
        {
            return obj is XLSheetArea other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Name.GetHashCode() * 397) ^ Area.GetHashCode();
            }
        }
    }
}
