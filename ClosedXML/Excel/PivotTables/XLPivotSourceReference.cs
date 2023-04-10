using ClosedXML.Excel.CalcEngine;
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotSourceReference : IEquatable<XLPivotSourceReference>
    {
        public XLPivotSourceReference(IXLRange range)
        {
            SourceType = XLPivotTableSourceType.Range;
            SourceRange = range;
        }

        public XLPivotSourceReference(IXLTable table)
        {
            SourceType = XLPivotTableSourceType.Table;
            SourceRange = table;
        }

        public IXLRange SourceRange { get; }

        public IXLTable? SourceTable => SourceRange as IXLTable;

        public XLPivotTableSourceType SourceType { get; }

        public override bool Equals(object obj)
        {
            var other = obj as XLPivotSourceReference;
            return Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (SourceRange.GetHashCode() * 397) ^ (int)SourceType;
            }
        }

        public bool Equals(XLPivotSourceReference? other)
        {
            if (other is null || SourceType != other.SourceType)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            switch (SourceType)
            {
                case XLPivotTableSourceType.Table:
                    return StringComparer.OrdinalIgnoreCase.Equals(SourceTable!.Name, other.SourceTable!.Name);

                case XLPivotTableSourceType.Range:
                    return XLRangeAddressComparer.IgnoreFixed.Equals(SourceRange.RangeAddress, other.SourceRange.RangeAddress);

                default:
                    throw new NotSupportedException();
            }
        }
    }
}
