using System;
namespace ClosedXML.Excel
{
    public interface IXLFilteredColumn
    {
        IXLFilteredColumn AddFilter<T>(T value) where T : IComparable<T>;
    }
}