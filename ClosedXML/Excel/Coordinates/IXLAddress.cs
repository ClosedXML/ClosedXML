using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>
    {
        IXLWorksheet Worksheet { get; }
        Int32 RowNumber { get; }
        Int32 ColumnNumber { get; }
        String ColumnLetter { get; }
        Boolean FixedRow { get; }
        Boolean FixedColumn { get; }
        String ToStringRelative();
        String ToStringRelative(Boolean includeSheet);
        String ToStringFixed();
        String ToStringFixed(XLReferenceStyle referenceStyle);
        String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet);
        String ToString(XLReferenceStyle referenceStyle);
        String UniqueId { get; }
    }
}
