using System;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A name in a worksheet. Unlike <see cref="IXLDefinedName"/>, this is basically only a reference.
    /// The actual 
    /// </summary>
    internal readonly struct XLName : IEquatable<XLName>
    {
        /// <summary>
        /// Name of a sheet. If null, the scope is a workbook. The sheet might not exist, e.g. it
        /// is only in a formula. The name of a sheet is not escaped.
        /// </summary>
        public string? SheetName { get; }

        /// <summary>
        /// The defined name in the scope. Case insensitive during comparisons.
        /// </summary>
        public string Name { get; }

        public XLName(string sheetName, string name)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(sheetName));

            if (name.Any(char.IsWhiteSpace))
                throw new ArgumentException("Name can't contain whitespace.");

            SheetName = sheetName;
            Name = name;
        }

        public XLName(string name)
        {
            if (name.Any(char.IsWhiteSpace))
                throw new ArgumentException("Name can't contain whitespace.");

            SheetName = null;
            Name = name;
        }

        public bool Equals(XLName other)
        {
            var differentScope = SheetName is null ^ other.SheetName is null;
            if (differentScope)
                return false;

            var bothWorkbookScope = SheetName is null && other.SheetName is null;
            if (bothWorkbookScope)
                return XLHelper.NameComparer.Equals(Name, other.Name);

            return XLHelper.NameComparer.Equals(Name, other.Name) &&
                   XLHelper.SheetComparer.Equals(SheetName, other.SheetName);
        }

        public override bool Equals(object? obj)
        {
            return obj is XLName other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (SheetName is not null ? XLHelper.SheetComparer.GetHashCode(SheetName) : 0) * 397;
                hashCode ^= XLHelper.NameComparer.GetHashCode(Name);
                return hashCode;
            }
        }

        public override string ToString()
        {
            var isWorkbookScoped = SheetName is null;
            return isWorkbookScoped ? Name : $"{SheetName}!{Name}";
        }
    }
}
