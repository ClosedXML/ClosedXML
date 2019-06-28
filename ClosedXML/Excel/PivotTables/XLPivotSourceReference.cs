using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.Patterns;
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotSourceReference : IXLPivotSourceReference
    {
        private IXLRange sourceRange;

        public IXLRange SourceRange
        {
            get { return sourceRange; }
            set
            {
                if (value is IXLTable)
                    SourceType = XLPivotTableSourceType.Table;
                else
                    SourceType = XLPivotTableSourceType.Range;

                sourceRange = value;
            }
        }

        public IXLTable SourceTable
        {
            get { return SourceRange as IXLTable; }
            set { SourceRange = value; }
        }

        public XLPivotTableSourceType SourceType { get; private set; }

        #region IEquatable interface

        public bool Equals(IXLPivotSourceReference other)
        {
            if (this.SourceType != other.SourceType) return false;

            switch (this.SourceType)
            {
                case XLPivotTableSourceType.Table:
                    return ClosedXMLValueComparer.DefaultComparer.Compare(this.SourceTable.Name, other.SourceTable.Name) == 0;

                case XLPivotTableSourceType.Range:
                    var rangeAddressComparer = new XLRangeAddressComparer(true);
                    return rangeAddressComparer.Equals(this.SourceRange.RangeAddress, other.SourceRange.RangeAddress);

                default:
                    throw new NotImplementedException();
            }
        }

        #endregion IEquatable interface
    }
}
