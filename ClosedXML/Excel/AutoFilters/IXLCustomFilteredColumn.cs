using System;
namespace ClosedXML.Excel
{
    public interface IXLCustomFilteredColumn
    {
        void EqualTo(XLCellValue value, bool reapply = true);
        void NotEqualTo(XLCellValue value, bool reapply = true);
        void GreaterThan(XLCellValue value, bool reapply = true);
        void LessThan(XLCellValue value, bool reapply = true);
        void EqualOrGreaterThan(XLCellValue value, bool reapply = true);
        void EqualOrLessThan(XLCellValue value, bool reapply = true);
        void BeginsWith(String value, bool reapply = true);
        void NotBeginsWith(String value, bool reapply = true);
        void EndsWith(String value, bool reapply = true);
        void NotEndsWith(String value, bool reapply = true);
        void Contains(String value, bool reapply = true);
        void NotContains(String value, bool reapply = true);
    }
}
