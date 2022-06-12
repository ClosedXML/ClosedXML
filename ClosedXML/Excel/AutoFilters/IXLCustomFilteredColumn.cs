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
        void BeginsWith(string value);
        void NotBeginsWith(string value);
        void EndsWith(string value);
        void NotEndsWith(string value);
        void Contains(string value);
        void NotContains(string value);
    }
}