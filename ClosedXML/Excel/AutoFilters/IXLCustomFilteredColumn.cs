using System;
namespace ClosedXML.Excel
{
    public interface IXLCustomFilteredColumn
    {
        void EqualTo(XLCellValue value);
        void NotEqualTo(XLCellValue value);
        void GreaterThan(XLCellValue value);
        void LessThan(XLCellValue value);
        void EqualOrGreaterThan(XLCellValue value);
        void EqualOrLessThan(XLCellValue value);
        void BeginsWith(String value);
        void NotBeginsWith(String value);
        void EndsWith(String value);
        void NotEndsWith(String value);
        void Contains(String value);
        void NotContains(String value);
    }
}
