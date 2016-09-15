﻿using System;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLFilterColumn : IXLFilterColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly Int32 _column;

        public XLFilterColumn(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        #region IXLFilterColumn Members

        public void Clear()
        {
            if (_autoFilter.Filters.ContainsKey(_column))
                _autoFilter.Filters.Remove(_column);
        }

        public IXLFilteredColumn AddFilter<T>(T value) where T: IComparable<T>
        {
            if (typeof(T) == typeof(String))
            {
                ApplyCustomFilter(value, XLFilterOperator.Equal,
                                  v =>
                                  v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase),
                                  XLFilterType.Regular);
            }
            else
            {
                ApplyCustomFilter(value, XLFilterOperator.Equal,
                                  v => v.CastTo<T>().CompareTo(value) == 0, XLFilterType.Regular);
            }
            return new XLFilteredColumn(_autoFilter, _column);
        }

        public void Top(Int32 value, XLTopBottomType type = XLTopBottomType.Items)
        {
            _autoFilter.Column(_column).TopBottomPart = XLTopBottomPart.Top;
            SetTopBottom(value, type);
        }

        public void Bottom(Int32 value, XLTopBottomType type = XLTopBottomType.Items)
        {
            _autoFilter.Column(_column).TopBottomPart = XLTopBottomPart.Bottom;
            SetTopBottom(value, type, false);
        }

        public void AboveAverage()
        {
            ShowAverage(true);
        }

        public void BelowAverage()
        {
            ShowAverage(false);
        }

        public IXLFilterConnector EqualTo<T>(T value) where T: IComparable<T>
        {
            if (typeof(T) == typeof(String))
            {
                return ApplyCustomFilter(value, XLFilterOperator.Equal,
                                         v =>
                                         v.ToString().Equals(value.ToString(),
                                                             StringComparison.InvariantCultureIgnoreCase));
            }

            return ApplyCustomFilter(value, XLFilterOperator.Equal,
                                     v => v.CastTo<T>().CompareTo(value) == 0);
        }

        public IXLFilterConnector NotEqualTo<T>(T value) where T: IComparable<T>
        {
            if (typeof(T) == typeof(String))
            {
                return ApplyCustomFilter(value, XLFilterOperator.NotEqual,
                                         v =>
                                         !v.ToString().Equals(value.ToString(),
                                                              StringComparison.InvariantCultureIgnoreCase));
            }
            
            return ApplyCustomFilter(value, XLFilterOperator.NotEqual,
                                        v => v.CastTo<T>().CompareTo(value) != 0);
        }

        public IXLFilterConnector GreaterThan<T>(T value) where T: IComparable<T>
        {
            return ApplyCustomFilter(value, XLFilterOperator.GreaterThan,
                                     v => v.CastTo<T>().CompareTo(value) > 0);
        }

        public IXLFilterConnector LessThan<T>(T value) where T: IComparable<T>
        {
            return ApplyCustomFilter(value, XLFilterOperator.LessThan,
                                     v => v.CastTo<T>().CompareTo(value) < 0);
        }

        public IXLFilterConnector EqualOrGreaterThan<T>(T value) where T: IComparable<T>
        {
            return ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan,
                                     v => v.CastTo<T>().CompareTo(value) >= 0);
        }

        public IXLFilterConnector EqualOrLessThan<T>(T value) where T: IComparable<T>
        {
            return ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan,
                                     v => v.CastTo<T>().CompareTo(value) <= 0);
        }

        public void Between<T>(T minValue, T maxValue) where T: IComparable<T>
        {
            EqualOrGreaterThan(minValue).And.EqualOrLessThan(maxValue);
        }

        public void NotBetween<T>(T minValue, T maxValue) where T: IComparable<T>
        {
            LessThan(minValue).Or.GreaterThan(maxValue);
        }

        public IXLFilterConnector BeginsWith(String value)
        {
            return ApplyCustomFilter(value + "*", XLFilterOperator.Equal,
                                     s => ((string)s).StartsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector NotBeginsWith(String value)
        {
            return ApplyCustomFilter(value + "*", XLFilterOperator.NotEqual,
                                     s => !((string)s).StartsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector EndsWith(String value)
        {
            return ApplyCustomFilter("*" + value, XLFilterOperator.Equal,
                                     s => ((string)s).EndsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector NotEndsWith(String value)
        {
            return ApplyCustomFilter("*" + value, XLFilterOperator.NotEqual,
                                     s => !((string)s).EndsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector Contains(String value)
        {
            return ApplyCustomFilter("*" + value + "*", XLFilterOperator.Equal,
                                     s => ((string)s).ToLower().Contains(value.ToLower()));
        }

        public IXLFilterConnector NotContains(String value)
        {
            return ApplyCustomFilter("*" + value + "*", XLFilterOperator.Equal,
                                     s => !((string)s).ToLower().Contains(value.ToLower()));
        }

        public XLFilterType FilterType { get; set; }

        public Int32 TopBottomValue { get; set; }
        public XLTopBottomType TopBottomType { get; set; }
        public XLTopBottomPart TopBottomPart { get; set; }

        public XLFilterDynamicType DynamicType { get; set; }
        public Double DynamicValue { get; set; }

        #endregion

        private void SetTopBottom(Int32 value, XLTopBottomType type, Boolean takeTop = true)
        {
            _autoFilter.Enabled = true;
            _autoFilter.Column(_column).SetFilterType(XLFilterType.TopBottom)
                                       .SetTopBottomValue(value)
                                       .SetTopBottomType(type);

            var values = GetValues(value, type, takeTop);

            Clear();
            _autoFilter.Filters.Add(_column, new List<XLFilter>());

            Boolean addToList = true;
            var ws = _autoFilter.Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();
            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());
                foreach (IXLRangeRow row in rows)
                {
                    Boolean foundOne = false;
                    foreach (double val in values)
                    {
                        Func<Object, Boolean> condition = v => (v as IComparable).CompareTo(val) == 0;
                        if (addToList)
                        {
                            _autoFilter.Filters[_column].Add(new XLFilter
                                                                 {
                                                                     Value = val,
                                                                     Operator = XLFilterOperator.Equal,
                                                                     Connector = XLConnector.Or,
                                                                     Condition = condition
                                                                 });
                        }

                        var cell = row.Cell(_column);
                        if (cell.DataType != XLCellValues.Number || !condition(cell.GetDouble())) continue;
                        row.WorksheetRow().Unhide();
                        foundOne = true;
                    }
                    if (!foundOne)
                        row.WorksheetRow().Hide();

                    addToList = false;
                }
            ws.ResumeEvents();
        }

        private IEnumerable<double> GetValues(int value, XLTopBottomType type, bool takeTop)
        {
            using (var column = _autoFilter.Range.Column(_column))
            {
                using (var subColumn = column.Column(2, column.CellCount()))
                {
                    var cellsUsed = subColumn.CellsUsed(c => c.DataType == XLCellValues.Number);
                    if (takeTop)
                    {
                        if (type == XLTopBottomType.Items)
                        {
                            return cellsUsed.Select(c => c.GetDouble()).OrderByDescending(d => d).Take(value).Distinct();
                        }
                        var numerics1 = cellsUsed.Select(c => c.GetDouble());
                        Int32 valsToTake1 = numerics1.Count() * value / 100;
                        return numerics1.OrderByDescending(d => d).Take(valsToTake1).Distinct();
                    }

                    if (type == XLTopBottomType.Items)
                    {
                        return cellsUsed.Select(c => c.GetDouble()).OrderBy(d => d).Take(value).Distinct();
                    }

                    var numerics = cellsUsed.Select(c => c.GetDouble());
                    Int32 valsToTake = numerics.Count() * value / 100;
                    return numerics.OrderBy(d => d).Take(valsToTake).Distinct();
                }
            }
        }

        private void ShowAverage(Boolean aboveAverage)
        {
            _autoFilter.Enabled = true;
            _autoFilter.Column(_column).SetFilterType(XLFilterType.Dynamic)
                .SetDynamicType(aboveAverage
                                    ? XLFilterDynamicType.AboveAverage
                                    : XLFilterDynamicType.BelowAverage);
            var values = GetAverageValues(aboveAverage);


            Clear();
            _autoFilter.Filters.Add(_column, new List<XLFilter>());

            Boolean addToList = true;
            var ws = _autoFilter.Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();
            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());
            
                foreach (IXLRangeRow row in rows)
                {
                    Boolean foundOne = false;
                    foreach (double val in values)
                    {
                        Func<Object, Boolean> condition = v => (v as IComparable).CompareTo(val) == 0;
                        if (addToList)
                        {
                            _autoFilter.Filters[_column].Add(new XLFilter
                                                                 {
                                                                     Value = val,
                                                                     Operator = XLFilterOperator.Equal,
                                                                     Connector = XLConnector.Or,
                                                                     Condition = condition
                                                                 });
                        }

                        var cell = row.Cell(_column);
                        if (cell.DataType != XLCellValues.Number || !condition(cell.GetDouble())) continue;
                        row.WorksheetRow().Unhide();
                        foundOne = true;
                    }

                    if (!foundOne)
                        row.WorksheetRow().Hide();

                    addToList = false;
                }
            
            ws.ResumeEvents();
        }

        private IEnumerable<double> GetAverageValues(bool aboveAverage)
        {
            using (var column = _autoFilter.Range.Column(_column))
            {
                using (var subColumn = column.Column(2, column.CellCount()))
                {
                    Double average = subColumn.CellsUsed(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble()).Average();

                    if (aboveAverage)
                    {
                        return
                            subColumn.CellsUsed(c => c.DataType == XLCellValues.Number).
                                Select(c => c.GetDouble()).Where(c => c > average).Distinct();
                    }

                    return
                        subColumn.CellsUsed(c => c.DataType == XLCellValues.Number).
                            Select(c => c.GetDouble()).Where(c => c < average).Distinct();

                }
            }
        }

        private IXLFilterConnector ApplyCustomFilter<T>(T value, XLFilterOperator op, Func<Object, Boolean> condition,
                                                        XLFilterType filterType = XLFilterType.Custom)
            where T: IComparable<T>
        {
            _autoFilter.Enabled = true;
            if (filterType == XLFilterType.Custom)
            {
                Clear();
                _autoFilter.Filters.Add(_column,
                                        new List<XLFilter>
                                            {
                                                new XLFilter
                                                    {
                                                        Value = value,
                                                        Operator = op,
                                                        Connector = XLConnector.Or,
                                                        Condition = condition
                                                    }
                                            });
            }
            else
            {
                List<XLFilter> filterList;
                if (_autoFilter.Filters.TryGetValue(_column, out filterList))
                    filterList.Add(new XLFilter
                                       {Value = value, Operator = op, Connector = XLConnector.Or, Condition = condition});
                else
                {
                    _autoFilter.Filters.Add(_column,
                                            new List<XLFilter>
                                                {
                                                    new XLFilter
                                                        {
                                                            Value = value,
                                                            Operator = op,
                                                            Connector = XLConnector.Or,
                                                            Condition = condition
                                                        }
                                                });
                }
            }
            _autoFilter.Column(_column).FilterType = filterType;
            Boolean isText = typeof(T) == typeof(String);
            var ws = _autoFilter.Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();
            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());
            foreach (IXLRangeRow row in rows)
            {
                Boolean match = isText
                                    ? condition(row.Cell(_column).GetString())
                                    : row.Cell(_column).DataType == XLCellValues.Number &&
                                        condition(row.Cell(_column).GetDouble());
                if (match)
                    row.WorksheetRow().Unhide();
                else
                    row.WorksheetRow().Hide();
            }
            ws.ResumeEvents();
            return new XLFilterConnector(_autoFilter, _column);
        }

        public IXLFilterColumn SetFilterType(XLFilterType value) { FilterType = value; return this; }
        public IXLFilterColumn SetTopBottomValue(Int32 value) { TopBottomValue = value; return this; }
        public IXLFilterColumn SetTopBottomType(XLTopBottomType value) { TopBottomType = value; return this; }
        public IXLFilterColumn SetTopBottomPart(XLTopBottomPart value) { TopBottomPart = value; return this; }
        public IXLFilterColumn SetDynamicType(XLFilterDynamicType value) { DynamicType = value; return this; }
        public IXLFilterColumn SetDynamicValue(Double value) { DynamicValue = value; return this; }

    }
}