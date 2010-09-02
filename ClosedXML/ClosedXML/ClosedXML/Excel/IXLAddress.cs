using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLAddress : IEqualityComparer<XLAddress>, IEquatable<XLAddress>, IComparable, IComparable<XLAddress>
    {
        Int32 Row { get; }
        Int32 Column { get; }
        String ColumnLetter { get; }
    }

    public static class IXLAddressMethods
    { 

    }
}
