using System;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    public interface IXLFilteredColumn
    {
        IXLFilteredColumn AddFilter<T>(T value) where T : IComparable<T>;
    }
}