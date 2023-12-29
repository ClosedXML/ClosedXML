using System;

namespace ClosedXML.Excel
{
    internal class XLCustomFilteredColumn : IXLCustomFilteredColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly Int32 _column;
        private readonly XLConnector _connector;

        public XLCustomFilteredColumn(XLAutoFilter autoFilter, Int32 column, XLConnector connector)
        {
            _autoFilter = autoFilter;
            _column = column;
            _connector = connector;
        }

        public void EqualTo(XLCellValue value)
        {
            ApplyCustomFilter(value, XLFilterOperator.Equal);
        }

        public void NotEqualTo(XLCellValue value)
        {
            ApplyCustomFilter(value, XLFilterOperator.NotEqual);
        }

        public void GreaterThan(XLCellValue value)
        {
            ApplyCustomFilter(value, XLFilterOperator.GreaterThan);
        }

        public void LessThan(XLCellValue value)
        {
            ApplyCustomFilter(value, XLFilterOperator.LessThan);
        }

        public void EqualOrGreaterThan(XLCellValue value)
        {
            ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan);
        }

        public void EqualOrLessThan(XLCellValue value)
        {
            ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan);
        }

        public void BeginsWith(String value)
        {
            ApplyWildcardCustomFilter(value + "*", true);
        }

        public void NotBeginsWith(String value)
        {
            ApplyWildcardCustomFilter(value + "*", false);
        }

        public void EndsWith(String value)
        {
            ApplyWildcardCustomFilter("*" + value, true);
        }

        public void NotEndsWith(String value)
        {
            ApplyWildcardCustomFilter("*" + value, false);
        }

        public void Contains(String value)
        {
            ApplyWildcardCustomFilter("*" + value + "*", true);
        }

        public void NotContains(String value)
        {
            ApplyWildcardCustomFilter("*" + value + "*", false);
        }

        private void ApplyCustomFilter(XLCellValue value, XLFilterOperator op)
        {
            _autoFilter.AddFilter(_column, XLFilter.CreateCustomFilter(value, op, _connector));
            _autoFilter.Reapply();
        }

        private void ApplyWildcardCustomFilter(string pattern, bool match)
        {
            _autoFilter.AddFilter(_column, XLFilter.CreateWildcardFilter(pattern, match, _connector));
            _autoFilter.Reapply();
        }
    }
}
