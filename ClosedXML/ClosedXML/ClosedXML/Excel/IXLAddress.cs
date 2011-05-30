using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>, IComparable, IComparable<IXLAddress>
    {
        IXLWorksheet Worksheet { get; }
        Int32 RowNumber { get; }
        Int32 ColumnNumber { get; }
        String ColumnLetter { get; }
        Boolean FixedRow { get; }
        Boolean FixedColumn { get; }
        String ToStringRelative();
        String ToStringFixed();
        String ToString(XLReferenceStyle referenceStyle);
    }
}
