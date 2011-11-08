using System;
using System.Linq;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLCustomFilteredColumn: IXLCustomFilteredColumn
    {
        private XLAutoFilter _autoFilter;
        private Int32 _column;
        private XLConnector _connector;
        public XLCustomFilteredColumn(XLAutoFilter autoFilter, Int32 column, XLConnector connector)
        {
            _autoFilter = autoFilter;
            _column = column;
            _connector = connector;
        }

        public void EqualTo<T>(T value) where T : IComparable<T>
        {
            if (typeof(T) == typeof(String))
                ApplyCustomFilter<T>(value, XLFilterOperator.Equal, v => v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
            else
                ApplyCustomFilter<T>(value, XLFilterOperator.Equal, v => (v.CastTo<T>() as IComparable).CompareTo(value) == 0);
        }

        public void NotEqualTo<T>(T value) where T : IComparable<T>
        {
            if (typeof(T) == typeof(String))
                ApplyCustomFilter<T>(value, XLFilterOperator.NotEqual, v => !v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
            else
                ApplyCustomFilter<T>(value, XLFilterOperator.NotEqual, v => (v.CastTo<T>() as IComparable).CompareTo(value) != 0);
        }

        public void GreaterThan<T>(T value) where T : IComparable<T>
        {
            ApplyCustomFilter<T>(value, XLFilterOperator.GreaterThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) > 0);
        }

        public void LessThan<T>(T value) where T : IComparable<T>
        {
            ApplyCustomFilter<T>(value, XLFilterOperator.LessThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) < 0);
        }

        public void EqualOrGreaterThan<T>(T value) where T : IComparable<T>
        {
            ApplyCustomFilter<T>(value, XLFilterOperator.EqualOrGreaterThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) >= 0);
        }

        public void EqualOrLessThan<T>(T value) where T : IComparable<T>
        {
            ApplyCustomFilter<T>(value, XLFilterOperator.EqualOrLessThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) <= 0);
        }

        public void BeginsWith(String value)
        {
            ApplyCustomFilter<String>(value.ToString() + "*", XLFilterOperator.Equal, s => ((string)s).StartsWith(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
        }

        public void NotBeginsWith(String value)
        {
            ApplyCustomFilter<String>(value.ToString() + "*", XLFilterOperator.NotEqual, s => !((string)s).StartsWith(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
        }

        public void EndsWith(String value)
        {
            ApplyCustomFilter<String>("*" + value.ToString(), XLFilterOperator.Equal, s => ((string)s).EndsWith(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
        }

        public void NotEndsWith(String value)
        {
            ApplyCustomFilter<String>("*" + value.ToString(), XLFilterOperator.NotEqual, s => !((string)s).EndsWith(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
        }

        public void Contains(String value)
        {
            ApplyCustomFilter<String>("*" + value.ToString() + "*", XLFilterOperator.Equal, s => ((string)s).ToLower().Contains(value.ToString().ToLower()));
        }
        public void NotContains(String value)
        {
            ApplyCustomFilter<String>("*" + value.ToString() + "*", XLFilterOperator.Equal, s => !((string)s).ToLower().Contains(value.ToString().ToLower()));
        }

        private void ApplyCustomFilter<T>(T value, XLFilterOperator op, Func<Object, Boolean> condition) where T : IComparable<T>
        {
            _autoFilter.Filters[_column].Add(new XLFilter { Value = value, Operator = op, Connector = _connector, Condition = condition });
            foreach (var row in _autoFilter.Range.Rows().Where(r => r.RowNumber() > 1))
            {
                if (_connector == XLConnector.And)
                {
                    if (!row.WorksheetRow().IsHidden)
                    {
                        if (condition(row.Cell(_column).GetValue<T>()))
                            row.WorksheetRow().Unhide();
                        else
                            row.WorksheetRow().Hide();
                    }
                }
                else if (condition(row.Cell(_column).GetValue<T>()))
                    row.WorksheetRow().Unhide();
            }
        }
    }
}