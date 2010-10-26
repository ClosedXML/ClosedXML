using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>, IComparable, IComparable<IXLAddress>
    {
        Int32 RowNumber { get; }
        Int32 ColumnNumber { get; }
        String ColumnLetter { get; }
    }

    public static class IXLAddressMethods
    { 

    }
}
