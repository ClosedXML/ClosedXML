using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>
    {
        string ColumnLetter { get; }
        int ColumnNumber { get; }
        bool FixedColumn { get; }
        bool FixedRow { get; }
        int RowNumber { get; }
        string UniqueId { get; }
        IXLWorksheet Worksheet { get; }

        string ToString(XLReferenceStyle referenceStyle);

        string ToString(XLReferenceStyle referenceStyle, bool includeSheet);

        string ToStringFixed();

        string ToStringFixed(XLReferenceStyle referenceStyle);

        string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet);

        string ToStringRelative();


        string ToStringRelative(bool includeSheet);
    }
}
