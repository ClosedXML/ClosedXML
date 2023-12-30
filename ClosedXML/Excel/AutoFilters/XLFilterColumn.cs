using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLFilterColumn : IXLFilterColumn, IXLFilteredColumn, IXLDateTimeGroupFilteredColumn, IEnumerable<XLFilter>
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly Int32 _column;
        private readonly List<XLFilter> _filters = new();
        public XLFilterColumn(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        #region IXLFilterColumn Members

        public void Clear()
        {
            _filters.Clear();
            FilterType = XLFilterType.None;
        }

        public IXLFilteredColumn AddFilter(XLCellValue value)
        {
            SwitchFilter(XLFilterType.Regular);
            AddFilter(XLFilter.CreateRegularFilter(value));
            _autoFilter.Reapply();
            return this;
        }

        public IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            SwitchFilter(XLFilterType.DateTimeGrouping);
            AddFilter(XLFilter.CreateDateGroupFilter(date, dateTimeGrouping));
            _autoFilter.Reapply();
            return this;
        }

        public void Top(Int32 value, XLTopBottomType type = XLTopBottomType.Items)
        {
            SetTopBottom(value, type, true);
        }

        public void Bottom(Int32 value, XLTopBottomType type = XLTopBottomType.Items)
        {
            SetTopBottom(value, type, false);
        }

        public void AboveAverage()
        {
            SetAverage(true);
        }

        public void BelowAverage()
        {
            SetAverage(false);
        }

        public IXLFilterConnector EqualTo(XLCellValue value)
        {
            return AddCustomFilter(value, XLFilterOperator.Equal);
        }

        public IXLFilterConnector NotEqualTo(XLCellValue value)
        {
            return AddCustomFilter(value, XLFilterOperator.NotEqual);
        }

        public IXLFilterConnector GreaterThan(XLCellValue value)
        {
            return AddCustomFilter(value, XLFilterOperator.GreaterThan);
        }

        public IXLFilterConnector LessThan(XLCellValue value)
        {
            return AddCustomFilter(value, XLFilterOperator.LessThan);
        }

        public IXLFilterConnector EqualOrGreaterThan(XLCellValue value)
        {
            return AddCustomFilter(value, XLFilterOperator.EqualOrGreaterThan);
        }

        public IXLFilterConnector EqualOrLessThan(XLCellValue value)
        {
            return AddCustomFilter(value, XLFilterOperator.EqualOrLessThan);
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
            return AddCustomFilter(value + "*", true);
        }

        public IXLFilterConnector NotBeginsWith(String value)
        {
            return AddCustomFilter(value + "*", false);
        }

        public IXLFilterConnector EndsWith(String value)
        {
            return AddCustomFilter("*" + value, true);
        }

        public IXLFilterConnector NotEndsWith(String value)
        {
            return AddCustomFilter("*" + value, false);
        }

        public IXLFilterConnector Contains(String value)
        {
            return AddCustomFilter("*" + value + "*", true);
        }

        public IXLFilterConnector NotContains(String value)
        {
            return AddCustomFilter("*" + value + "*", false);
        }

        public XLFilterType FilterType { get; set; }

        public Int32 TopBottomValue { get; set; }
        public XLTopBottomType TopBottomType { get; set; }
        public XLTopBottomPart TopBottomPart { get; set; }

        public XLFilterDynamicType DynamicType { get; set; }
        public Double DynamicValue { get; set; }

        #endregion IXLFilterColumn Members

        public IEnumerator<XLFilter> GetEnumerator() => _filters.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private void SetTopBottom(Int32 value, XLTopBottomType type, Boolean takeTop)
        {
            ResetFilter(XLFilterType.TopBottom);
            TopBottomValue = value;
            TopBottomType = type;
            TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

            var filterValue = GetTopBottomFilterValue(type, value, takeTop);
            AddFilter(XLFilter.CreateTopBottom(takeTop, filterValue));
            _autoFilter.Reapply();
        }

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

        private void SetAverage(Boolean aboveAverage)
        {
            ResetFilter(XLFilterType.Dynamic);
            DynamicType = aboveAverage
                ? XLFilterDynamicType.AboveAverage
                : XLFilterDynamicType.BelowAverage;
            var average = GetAverageFilterValue();
            AddFilter(XLFilter.CreateAverage(average, aboveAverage));
            _autoFilter.Reapply();

            double GetAverageFilterValue()
            {
                var column = _autoFilter.Range.Column(_column);
                var subColumn = column.Column(2, column.CellCount());
                return subColumn.CellsUsed(c => c.CachedValue.IsUnifiedNumber)
                    .Select(c => c.CachedValue.GetUnifiedNumber())
                    .DefaultIfEmpty(Double.NaN)
                    .Average();
            }
        }

        private IXLFilterConnector AddCustomFilter(XLCellValue value, XLFilterOperator op)
        {
            ResetFilter(XLFilterType.Custom);
            AddFilter(XLFilter.CreateCustomFilter(value, op, XLConnector.Or));
            _autoFilter.Reapply();
            return new XLFilterConnector(_autoFilter, _column);
        }

        private IXLFilterConnector AddCustomFilter(string pattern, bool match)
        {
            SwitchFilter(XLFilterType.Custom);
            AddFilter(XLFilter.CreateWildcardFilter(pattern, match, XLConnector.Or));
            _autoFilter.Reapply();
            return new XLFilterConnector(_autoFilter, _column);
        }

        private void ResetFilter(XLFilterType type)
        {
            Clear();
            _autoFilter.IsEnabled = true;
            FilterType = type;
        }

        private void SwitchFilter(XLFilterType type)
        {
            _autoFilter.IsEnabled = true;
            if (FilterType == type)
                return;

            Clear();
            FilterType = type;
        }

        internal void AddFilter(XLFilter filter)
        {
            _filters.Add(filter);
        }

        internal bool Check(IXLCell cell)
        {
            if (_filters.Count == 0)
                return true;

            if (_filters.Count == 1)
                return _filters[0].Condition(cell);

            // All filter conditions are connected by a single type of logical condition. Regular
            // filters use 'Or', custom has up to two clauses connected by 'And'/'Or' and rest is
            // single clause.
            var connector = _filters[0].Connector;
            return connector switch
            {
                XLConnector.And => _filters.All(filter => filter.Condition(cell)),
                XLConnector.Or  => _filters.Any(filter => filter.Condition(cell)),
                _ => throw new NotSupportedException(),
            };
        }
    }
}
