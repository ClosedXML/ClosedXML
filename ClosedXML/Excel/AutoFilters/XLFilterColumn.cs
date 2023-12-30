using System;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLFilterColumn : IXLFilterColumn, IXLFilteredColumn, IXLDateTimeGroupFilteredColumn
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
            return this;
        }

        public IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            _autoFilter.IsEnabled = true;
            FilterType = XLFilterType.DateTimeGrouping;
            _autoFilter.AddFilter(_column, XLFilter.CreateDateGroupFilter(date, dateTimeGrouping));
            _autoFilter.Reapply();

            return this;
        }

        public void Top(Int32 value, XLTopBottomType type = XLTopBottomType.Items)
        {
            _autoFilter.Column(_column).TopBottomPart = XLTopBottomPart.Top;
            SetTopBottom(value, type, true);
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

        public IXLFilterConnector BeginsWith(String value)
        {
            return ApplyCustomFilter(value + "*", true);
        }

        public IXLFilterConnector NotBeginsWith(String value)
        {
            return ApplyCustomFilter(value + "*", false);
        }

        public IXLFilterConnector EndsWith(String value)
        {
            return ApplyCustomFilter("*" + value, true);
        }

        public IXLFilterConnector NotEndsWith(String value)
        {
            return ApplyCustomFilter("*" + value, false);
        }

        public IXLFilterConnector Contains(String value)
        {
            return ApplyCustomFilter("*" + value + "*", true);
        }

        public IXLFilterConnector NotContains(String value)
        {
            return ApplyCustomFilter("*" + value + "*", false);
        }

        public XLFilterType FilterType { get; set; }

        public Int32 TopBottomValue { get; set; }
        public XLTopBottomType TopBottomType { get; set; }
        public XLTopBottomPart TopBottomPart { get; set; }

        public XLFilterDynamicType DynamicType { get; set; }
        public Double DynamicValue { get; set; }

        #endregion IXLFilterColumn Members

        private void SetTopBottom(Int32 value, XLTopBottomType type, Boolean takeTop)
        {
            _autoFilter.IsEnabled = true;
            Clear();
            _autoFilter.Column(_column).SetFilterType(XLFilterType.TopBottom)
                                       .SetTopBottomValue(value)
                                       .SetTopBottomType(type);

            var filterValue = GetTopBottomFilterValue(type, value, takeTop);
            _autoFilter.AddFilter(_column, XLFilter.CreateTopBottom(takeTop, filterValue));
            _autoFilter.Reapply();
        }

        /// <summary>
        /// Get a border value for top/bottom filter value.
        /// </summary>
        /// <param name="type">Content of <paramref name="value"/>.</param>
        /// <param name="value">Either percents or items.</param>
        /// <param name="takeTop">Take top (<c>true</c>) or bottom (<c>false</c>).</param>
        private double GetTopBottomFilterValue(XLTopBottomType type, int value, bool takeTop)
        {
            var column = _autoFilter.Range.Column(_column);
            var subColumn = column.Column(2, column.CellCount());
            var columnNumbers = subColumn.CellsUsed(c => c.CachedValue.IsUnifiedNumber).Select(c => c.CachedValue.GetUnifiedNumber());
            var comparer = takeTop
                ? Comparer<double>.Create((x, y) => -x.CompareTo(y))
                : Comparer<double>.Create((x, y) => x.CompareTo(y));

            switch (type)
            {
                case XLTopBottomType.Items:
                    var itemCount = value;
                    return columnNumbers.OrderBy(d => d, comparer).Take(itemCount).DefaultIfEmpty(double.NaN).LastOrDefault();
                case XLTopBottomType.Percent:
                    var percent = value;
                    var materializedNumbers = columnNumbers.ToArray();
                    var itemCountByPercents = materializedNumbers.Length * percent / 100;
                    return materializedNumbers.OrderBy(d => d, comparer).Take(itemCountByPercents).DefaultIfEmpty(Double.NaN).LastOrDefault();
                default:
                    throw new NotSupportedException();
            }
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

            foreach (double val in values)
            {
                Func<IXLCell, Boolean> condition = v => v.CachedValue.IsUnifiedNumber && v.CachedValue.GetUnifiedNumber().Equals(val);
                _autoFilter.AddFilter(_column, new XLFilter
                {
                    Value = val,
                    Operator = XLFilterOperator.Equal,
                    Connector = XLConnector.Or,
                    Condition = condition
                });
            }

            _autoFilter.Reapply();
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

        private IXLFilterConnector ApplyCustomFilter(XLCellValue value, XLFilterOperator op)
        {
            _autoFilter.IsEnabled = true;
            Clear();

            _autoFilter.AddFilter(_column, XLFilter.CreateCustomFilter(value, op, XLConnector.Or));
            _autoFilter.Column(_column).FilterType = XLFilterType.Custom;
            _autoFilter.Reapply();
            return new XLFilterConnector(_autoFilter, _column);
        }

        private IXLFilterConnector ApplyCustomFilter(string pattern, bool match)
        {
            _autoFilter.IsEnabled = true;
            Clear();

            _autoFilter.AddFilter(_column, XLFilter.CreateWildcardFilter(pattern, match, XLConnector.Or));
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
