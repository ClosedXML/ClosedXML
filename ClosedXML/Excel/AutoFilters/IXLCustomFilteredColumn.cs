using System;
namespace ClosedXML.Excel
{
    public interface IXLCustomFilteredColumn
    {
        void EqualTo<T>(T value) where T : IComparable<T>;
        void NotEqualTo<T>(T value) where T : IComparable<T>;
        void GreaterThan<T>(T value) where T : IComparable<T>;
        void LessThan<T>(T value) where T : IComparable<T>;
        void EqualOrGreaterThan<T>(T value) where T : IComparable<T>;
        void EqualOrLessThan<T>(T value) where T : IComparable<T>;
        void BeginsWith(String value);
        void NotBeginsWith(String value);
        void EndsWith(String value);
        void NotEndsWith(String value);
        void Contains(String value);
        void NotContains(String value);
    }
}