using System;

namespace ClosedXML.Excel
{
    internal class XLCustomFilteredColumn : IXLCustomFilteredColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly XLFilterColumn _filterColumn;
        private readonly XLConnector _connector;

        public XLCustomFilteredColumn(XLAutoFilter autoFilter, XLFilterColumn filterColumn, XLConnector connector)
        {
            _autoFilter = autoFilter;
            _filterColumn = filterColumn;
            _connector = connector;
        }

        public void EqualTo(XLCellValue value, Boolean reapply)
        {
            ApplyCustomFilter(value, XLFilterOperator.Equal, reapply);
        }

        public void NotEqualTo(XLCellValue value, Boolean reapply)
        {
            ApplyCustomFilter(value, XLFilterOperator.NotEqual, reapply);
        }

        public void GreaterThan(XLCellValue value, Boolean reapply)
        {
            ApplyCustomFilter(value, XLFilterOperator.GreaterThan, reapply);
        }

        public void LessThan(XLCellValue value, Boolean reapply)
        {
            ApplyCustomFilter(value, XLFilterOperator.LessThan, reapply);
        }

        public void EqualOrGreaterThan(XLCellValue value, Boolean reapply)
        {
            ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan, reapply);
        }

        public void EqualOrLessThan(XLCellValue value, Boolean reapply)
        {
            ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan, reapply);
        }

        public void BeginsWith(String value, Boolean reapply)
        {
            ApplyWildcardCustomFilter(value + "*", true, reapply);
        }

        public void NotBeginsWith(String value, Boolean reapply)
        {
            ApplyWildcardCustomFilter(value + "*", false, reapply);
        }

        public void EndsWith(String value, Boolean reapply)
        {
            ApplyWildcardCustomFilter("*" + value, true, reapply);
        }

        public void NotEndsWith(String value, Boolean reapply)
        {
            ApplyWildcardCustomFilter("*" + value, false, reapply);
        }

        public void Contains(String value, Boolean reapply)
        {
            ApplyWildcardCustomFilter("*" + value + "*", true, reapply);
        }

        public void NotContains(String value, Boolean reapply)
        {
            ApplyWildcardCustomFilter("*" + value + "*", false, reapply);
        }

        private void ApplyCustomFilter(XLCellValue value, XLFilterOperator op, bool reapply)
        {
            _filterColumn.AddFilter(XLFilter.CreateCustomFilter(value, op, _connector));
            if (reapply)
                _autoFilter.Reapply();
        }

        private void ApplyWildcardCustomFilter(string pattern, bool match, bool reapply)
        {
            _filterColumn.AddFilter(XLFilter.CreateWildcardFilter(pattern, match, _connector));
            if (reapply)
                _autoFilter.Reapply();
        }
    }
}
