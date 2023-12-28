using System;
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

        public IXLFilteredColumn AddFilter(XLCellValue value)
        {
            // TODO: If different filter type, clear them
            _autoFilter.IsEnabled = true;
            FilterType = XLFilterType.Regular;
            _autoFilter.AddFilter(_column, XLFilter.CreateRegularFilter(value));
            _autoFilter.Reapply();
            return new XLFilteredColumn(_autoFilter, _column);
        }

        public IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            _autoFilter.IsEnabled = true;
            FilterType = XLFilterType.DateTimeGrouping;
            _autoFilter.AddFilter(_column, XLFilter.CreateRegularDateGroupFilter(date, dateTimeGrouping));
            _autoFilter.Reapply();

            return new XLDateTimeGroupFilteredColumn(_autoFilter, _column);
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

        public IXLFilterConnector EqualTo(XLCellValue value)
        {
            return ApplyCustomFilter(value, XLFilterOperator.Equal);
        }

        public IXLFilterConnector NotEqualTo(XLCellValue value)
        {
            return ApplyCustomFilter(value, XLFilterOperator.NotEqual);
        }

        public IXLFilterConnector GreaterThan(XLCellValue value)
        {
            return ApplyCustomFilter(value, XLFilterOperator.GreaterThan);
        }

        public IXLFilterConnector LessThan(XLCellValue value)
        {
            return ApplyCustomFilter(value, XLFilterOperator.LessThan);
        }

        public IXLFilterConnector EqualOrGreaterThan(XLCellValue value)
        {
            return ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan);
        }

        public IXLFilterConnector EqualOrLessThan(XLCellValue value)
        {
            return ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan);
        }

        public void Between(XLCellValue minValue, XLCellValue maxValue)
        {
            EqualOrGreaterThan(minValue).And.EqualOrLessThan(maxValue);
        }

        public void NotBetween(XLCellValue minValue, XLCellValue maxValue)
        {
            LessThan(minValue).Or.GreaterThan(maxValue);
        }

        public static Func<String, Object, Boolean> BeginsWithFunction { get; } = (value, input) => ((string)input).StartsWith(value, StringComparison.InvariantCultureIgnoreCase);

        public IXLFilterConnector BeginsWith(String value)
        {
            return ApplyCustomFilter(value + "*", XLFilterOperator.Equal, s => BeginsWithFunction(value, s));
        }

        public IXLFilterConnector NotBeginsWith(String value)
        {
            return ApplyCustomFilter(value + "*", XLFilterOperator.NotEqual, s => !BeginsWithFunction(value, s));
        }

        public static Func<String, Object, Boolean> EndsWithFunction { get; } = (value, input) => ((string)input).EndsWith(value, StringComparison.InvariantCultureIgnoreCase);

        public IXLFilterConnector EndsWith(String value)
        {
            return ApplyCustomFilter("*" + value, XLFilterOperator.Equal, s => EndsWithFunction(value, s));
        }

        public IXLFilterConnector NotEndsWith(String value)
        {
            return ApplyCustomFilter("*" + value, XLFilterOperator.NotEqual, s => !EndsWithFunction(value, s));
        }

        public static Func<String, Object, Boolean> ContainsFunction { get; } = (value, input) => ((string)input).IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;

        public IXLFilterConnector Contains(String value)
        {
            return ApplyCustomFilter("*" + value + "*", XLFilterOperator.Equal, s => ContainsFunction(value, s));
        }

        public IXLFilterConnector NotContains(String value)
        {
            return ApplyCustomFilter("*" + value + "*", XLFilterOperator.NotEqual, s => !ContainsFunction(value, s));
        }

        public XLFilterType FilterType { get; set; }

        public Int32 TopBottomValue { get; set; }
        public XLTopBottomType TopBottomType { get; set; }
        public XLTopBottomPart TopBottomPart { get; set; }

        public XLFilterDynamicType DynamicType { get; set; }
        public Double DynamicValue { get; set; }

        #endregion IXLFilterColumn Members

        private void SetTopBottom(Int32 value, XLTopBottomType type, Boolean takeTop = true)
        {
            _autoFilter.IsEnabled = true;
            _autoFilter.Column(_column).SetFilterType(XLFilterType.TopBottom)
                                       .SetTopBottomValue(value)
                                       .SetTopBottomType(type);

            var values = GetValues(value, type, takeTop).ToArray();

            Clear();

            Boolean addToList = true;
            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());
            foreach (IXLRangeRow row in rows)
            {
                Boolean foundOne = false;
                foreach (double val in values)
                {
                    Func<Object, Boolean> condition = v => ((IComparable)v).CompareTo(val) == 0;
                    if (addToList)
                    {
                        _autoFilter.AddFilter(_column, new XLFilter
                        {
                            Value = val,
                            Operator = XLFilterOperator.Equal,
                            Connector = XLConnector.Or,
                            Condition = condition
                        });
                    }

                    var cell = row.Cell(_column);
                    if (cell.DataType != XLDataType.Number || !condition(cell.GetDouble())) continue;
                    row.WorksheetRow().Unhide();
                    foundOne = true;
                }
                if (!foundOne)
                    row.WorksheetRow().Hide();

                addToList = false;
            }
        }

        private IEnumerable<double> GetValues(int value, XLTopBottomType type, bool takeTop)
        {
            var column = _autoFilter.Range.Column(_column);
            var subColumn = column.Column(2, column.CellCount());
            var columnNumbers = subColumn.CellsUsed(c => c.DataType == XLDataType.Number).Select(c => c.GetDouble());
            var comparer = takeTop
                ? Comparer<double>.Create((x, y) => -x.CompareTo(y))
                : Comparer<double>.Create((x, y) => x.CompareTo(y));

            if (type == XLTopBottomType.Items)
            {
                var itemCount = value;
                return columnNumbers.OrderBy(d => d, comparer).Take(itemCount).Distinct();
            }

            var numerics = columnNumbers.ToArray();
            var percent = value;
            Int32 itemCountByPercents = numerics.Length * percent / 100;
            return numerics.OrderBy(d => d, comparer).Take(itemCountByPercents).Distinct();
        }

        private void ShowAverage(Boolean aboveAverage)
        {
            _autoFilter.IsEnabled = true;
            _autoFilter.Column(_column).SetFilterType(XLFilterType.Dynamic)
                .SetDynamicType(aboveAverage
                                    ? XLFilterDynamicType.AboveAverage
                                    : XLFilterDynamicType.BelowAverage);
            var values = GetAverageValues(aboveAverage).ToArray();

            Clear();

            Boolean addToList = true;
            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());

            foreach (IXLRangeRow row in rows)
            {
                Boolean foundOne = false;
                foreach (double val in values)
                {
                    Func<Object, Boolean> condition = v => ((IComparable)v).CompareTo(val) == 0;
                    if (addToList)
                    {
                        _autoFilter.AddFilter(_column, new XLFilter
                        {
                            Value = val,
                            Operator = XLFilterOperator.Equal,
                            Connector = XLConnector.Or,
                            Condition = condition
                        });
                    }

                    var cell = row.Cell(_column);
                    if (cell.DataType != XLDataType.Number || !condition(cell.GetDouble())) continue;
                    row.WorksheetRow().Unhide();
                    foundOne = true;
                }

                if (!foundOne)
                    row.WorksheetRow().Hide();

                addToList = false;
            }
        }

        private IEnumerable<double> GetAverageValues(bool aboveAverage)
        {
            var column = _autoFilter.Range.Column(_column);
            var subColumn = column.Column(2, column.CellCount());
            Double average = subColumn.CellsUsed(c => c.DataType == XLDataType.Number).Select(c => c.GetDouble())
                .Average();

            var distinctNumbers = subColumn
                .CellsUsed(c => c.DataType == XLDataType.Number)
                .Select(c => c.GetDouble())
                .Distinct();
            return aboveAverage
                ? distinctNumbers.Where(c => c > average)
                : distinctNumbers.Where(c => c < average);
        }

        private IXLFilterConnector ApplyCustomFilter<T>(T value, XLFilterOperator op, Func<Object, Boolean> condition)
            where T : IComparable<T>
        {
            _autoFilter.IsEnabled = true;
            Clear();

            _autoFilter.AddFilter(_column, new XLFilter
            {
                Value = value,
                Operator = op,
                Connector = XLConnector.Or,
                Condition = condition
            });
            _autoFilter.Column(_column).FilterType = XLFilterType.Custom;
            _autoFilter.Reapply();
            return new XLFilterConnector(_autoFilter, _column);
        }

        private IXLFilterConnector ApplyCustomFilter(XLCellValue value, XLFilterOperator op)
        {
            _autoFilter.IsEnabled = true;
            Clear();

            _autoFilter.AddFilter(_column, XLFilter.CreateCustomFilter(value, op, XLConnector.Or));
            _autoFilter.Column(_column).FilterType = XLFilterType.Custom;
            _autoFilter.Reapply();
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
