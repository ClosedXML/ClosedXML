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

        public IXLFilteredColumn AddFilter(XLCellValue value, bool reapply)
        {
            SwitchFilter(XLFilterType.Regular);
            AddFilter(XLFilter.CreateRegularFilter(value));
            if (reapply)
                _autoFilter.Reapply();

            return this;
        }

        public IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping, bool reapply)
        {
            SwitchFilter(XLFilterType.DateTimeGrouping);
            AddFilter(XLFilter.CreateDateGroupFilter(date, dateTimeGrouping));
            if (reapply)
                _autoFilter.Reapply();

            return this;
        }

        public void Top(Int32 value, XLTopBottomType type, bool reapply)
        {
            SetTopBottom(value, type, takeTop: true, reapply);
        }

        public void Bottom(Int32 value, XLTopBottomType type, bool reapply)
        {
            SetTopBottom(value, type, takeTop: false, reapply);
        }

        public void AboveAverage(bool reapply)
        {
            SetAverage(aboveAverage: true, reapply);
        }

        public void BelowAverage(bool reapply)
        {
            SetAverage(aboveAverage: false, reapply);
        }

        public IXLFilterConnector EqualTo(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value, XLFilterOperator.Equal, reapply);
        }

        public IXLFilterConnector NotEqualTo(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value, XLFilterOperator.NotEqual, reapply);
        }

        public IXLFilterConnector GreaterThan(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value, XLFilterOperator.GreaterThan, reapply);
        }

        public IXLFilterConnector LessThan(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value, XLFilterOperator.LessThan, reapply);
        }

        public IXLFilterConnector EqualOrGreaterThan(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value, XLFilterOperator.EqualOrGreaterThan, reapply);
        }

        public IXLFilterConnector EqualOrLessThan(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value, XLFilterOperator.EqualOrLessThan, reapply);
        }

        public void Between(XLCellValue minValue, XLCellValue maxValue, Boolean reapply)
        {
            EqualOrGreaterThan(minValue, false).And.EqualOrLessThan(maxValue, reapply);
        }

        public void NotBetween(XLCellValue minValue, XLCellValue maxValue, Boolean reapply)
        {
            LessThan(minValue, false).Or.GreaterThan(maxValue, reapply);
        }

        public IXLFilterConnector BeginsWith(String value, Boolean reapply)
        {
            return AddCustomFilter(value + "*", true, reapply);
        }

        public IXLFilterConnector NotBeginsWith(String value, Boolean reapply)
        {
            return AddCustomFilter(value + "*", false, reapply);
        }

        public IXLFilterConnector EndsWith(String value, Boolean reapply)
        {
            return AddCustomFilter("*" + value, true, reapply);
        }

        public IXLFilterConnector NotEndsWith(String value, Boolean reapply)
        {
            return AddCustomFilter("*" + value, false, reapply);
        }

        public IXLFilterConnector Contains(String value, Boolean reapply)
        {
            return AddCustomFilter("*" + value + "*", true, reapply);
        }

        public IXLFilterConnector NotContains(String value, Boolean reapply)
        {
            return AddCustomFilter("*" + value + "*", false, reapply);
        }

        public XLFilterType FilterType { get; set; }

        public Int32 TopBottomValue { get; set; }
        public XLTopBottomType TopBottomType { get; set; }
        public XLTopBottomPart TopBottomPart { get; set; }

        public XLFilterDynamicType DynamicType { get; set; }

        /// <summary>
        /// Basically average for dynamic filters. Value is refreshed during filter reapply.
        /// </summary>
        public Double DynamicValue { get; set; } = double.NaN;

        #endregion IXLFilterColumn Members

        public IEnumerator<XLFilter> GetEnumerator() => _filters.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private void SetTopBottom(Int32 value, XLTopBottomType type, Boolean takeTop, Boolean reapply)
        {
            ResetFilter(XLFilterType.TopBottom);
            TopBottomValue = value;
            TopBottomType = type;
            TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

            var filterValue = GetTopBottomFilterValue(type, value, takeTop);
            AddFilter(XLFilter.CreateTopBottom(takeTop, filterValue));
            if (reapply)
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

        private void SetAverage(Boolean aboveAverage, Boolean reapply)
        {
            ResetFilter(XLFilterType.Dynamic);
            DynamicType = aboveAverage
                ? XLFilterDynamicType.AboveAverage
                : XLFilterDynamicType.BelowAverage;

            if (reapply)
            {
                // `Average` is recalculated during reapply, so no need to calculate it twice.
                AddFilter(XLFilter.CreateAverage(double.NaN, aboveAverage));
                _autoFilter.Reapply();
            }
            else
            {
                // Calculate average, so it is saved to a workbook, even if filters are never reapplies.
                DynamicValue = GetAverageFilterValue();
                AddFilter(XLFilter.CreateAverage(DynamicValue, aboveAverage));
            }
        }

        private double GetAverageFilterValue()
        {
            var column = _autoFilter.Range.Column(_column);
            var subColumn = column.Column(2, column.CellCount());
            return subColumn.CellsUsed(c => c.CachedValue.IsUnifiedNumber)
                .Select(c => c.CachedValue.GetUnifiedNumber())
                .DefaultIfEmpty(Double.NaN)
                .Average();
        }

        private IXLFilterConnector AddCustomFilter(XLCellValue value, XLFilterOperator op, Boolean reapply)
        {
            ResetFilter(XLFilterType.Custom);
            AddFilter(XLFilter.CreateCustomFilter(value, op, XLConnector.Or));
            if (reapply)
                _autoFilter.Reapply();

            return new XLFilterConnector(_autoFilter, this);
        }

        private IXLFilterConnector AddCustomFilter(string pattern, bool match, bool reapply)
        {
            SwitchFilter(XLFilterType.Custom);
            AddFilter(XLFilter.CreateWildcardFilter(pattern, match, XLConnector.Or));
            if (reapply)
                _autoFilter.Reapply();

            return new XLFilterConnector(_autoFilter, this);
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

        internal void Refresh()
        {
            if (FilterType == XLFilterType.Dynamic && _filters.Count > 0)
            {
                // Update average value of a filter, so it is saved correctly and filter uses
                // correct value, even is cell values changed and avg was stale.
                DynamicValue = GetAverageFilterValue();
                _filters[0].Value = DynamicValue;
            }
        }

        internal bool Check(IXLCell cell)
        {
            if (_filters.Count == 0)
                return true;

            if (_filters.Count == 1)
                return _filters[0].Condition(cell, this);

            // All filter conditions are connected by a single type of logical condition. Regular
            // filters use 'Or', custom has up to two clauses connected by 'And'/'Or' and rest is
            // single clause.
            var connector = _filters[0].Connector;
            return connector switch
            {
                XLConnector.And => _filters.All(filter => filter.Condition(cell, this)),
                XLConnector.Or => _filters.Any(filter => filter.Condition(cell, this)),
                _ => throw new NotSupportedException(),
            };
        }
    }
}
