﻿using System;
using System.Linq;

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

        public void EqualTo<T>(T value) where T: IComparable<T>
        {
            if (typeof(T) == typeof(String))
            {
                ApplyCustomFilter(value, XLFilterOperator.Equal,
                                  v =>
                                  v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
            }
            else
            {
                ApplyCustomFilter(value, XLFilterOperator.Equal,
                                  v => v.CastTo<T>().CompareTo(value) == 0);
            }
        }

        public void NotEqualTo<T>(T value) where T: IComparable<T>
        {
            if (typeof(T) == typeof(String))
            {
                ApplyCustomFilter(value, XLFilterOperator.NotEqual,
                                  v =>
                                  !v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
            }
            else
            {
                ApplyCustomFilter(value, XLFilterOperator.NotEqual,
                                  v => v.CastTo<T>().CompareTo(value) != 0);
            }
        }

        public void GreaterThan<T>(T value) where T: IComparable<T>
        {
            ApplyCustomFilter(value, XLFilterOperator.GreaterThan,
                              v => v.CastTo<T>().CompareTo(value) > 0);
        }

        public void LessThan<T>(T value) where T: IComparable<T>
        {
            ApplyCustomFilter(value, XLFilterOperator.LessThan, v => v.CastTo<T>().CompareTo(value) < 0);
        }

        public void EqualOrGreaterThan<T>(T value) where T: IComparable<T>
        {
            ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan,
                              v => v.CastTo<T>().CompareTo(value) >= 0);
        }

        public void EqualOrLessThan<T>(T value) where T: IComparable<T>
        {
            ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan,
                              v => v.CastTo<T>().CompareTo(value) <= 0);
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
            where T: IComparable<T>
        {
            _autoFilter.Filters[_column].Add(new XLFilter
                                                 {
                                                     Value = value,
                                                     Operator = op,
                                                     Connector = _connector,
                                                     Condition = condition
                                                 });
            using (var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount()))
            {
                foreach (IXLRangeRow row in rows)
                {
                    if (_connector == XLConnector.And)
                    {
                        if (!row.WorksheetRow().IsHidden)
                        {
                            if (condition(row.Cell(_column).GetValue<T>()))
                                row.WorksheetRow().Unhide().Dispose();
                            else
                                row.WorksheetRow().Hide().Dispose();
                        }
                    }
                    else if (condition(row.Cell(_column).GetValue<T>()))
                        row.WorksheetRow().Unhide().Dispose();
                }
            }
        }
    }
}