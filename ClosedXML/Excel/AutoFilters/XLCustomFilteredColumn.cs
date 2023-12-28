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

        #region IXLCustomFilteredColumn Members

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
            ApplyCustomFilter(value + "*", XLFilterOperator.Equal,
                              s => ((string)s).StartsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public void NotBeginsWith(String value)
        {
            ApplyCustomFilter(value + "*", XLFilterOperator.NotEqual,
                              s =>
                              !((string)s).StartsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public void EndsWith(String value)
        {
            ApplyCustomFilter("*" + value, XLFilterOperator.Equal,
                              s => ((string)s).EndsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public void NotEndsWith(String value)
        {
            ApplyCustomFilter("*" + value, XLFilterOperator.NotEqual,
                              s => !((string)s).EndsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public void Contains(String value)
        {
            ApplyCustomFilter("*" + value + "*", XLFilterOperator.Equal,
                              s => ((string)s).ToLower().Contains(value.ToLower()));
        }

        public void NotContains(String value)
        {
            ApplyCustomFilter("*" + value + "*", XLFilterOperator.Equal,
                              s => !((string)s).ToLower().Contains(value.ToLower()));
        }

        #endregion

        private void ApplyCustomFilter<T>(T value, XLFilterOperator op, Func<Object, Boolean> condition)
            where T : IComparable<T>
        {
            _autoFilter.AddFilter(_column, new XLFilter
            {
                Value = value,
                Operator = op,
                Connector = _connector,
                Condition = condition
            });
            _autoFilter.Reapply();
        }

        private void ApplyCustomFilter(XLCellValue value, XLFilterOperator op)
        {
            _autoFilter.AddFilter(_column, XLFilter.CreateCustomFilter(value, op, _connector));
            _autoFilter.Reapply();
        }
    }
}
