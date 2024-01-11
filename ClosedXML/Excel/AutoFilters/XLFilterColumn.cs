using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLFilterColumn : IXLFilterColumn, IXLFilteredColumn, IEnumerable<XLFilter>
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

        public void Clear(bool reapply)
        {
            _filters.Clear();
            FilterType = XLFilterType.None;
            if (reapply)
                _autoFilter.Reapply();
        }

        public IXLFilteredColumn AddFilter(XLCellValue value, bool reapply)
        {
            SwitchFilter(XLFilterType.Regular);
            AddFilter(XLFilter.CreateRegularFilter(value.ToString()), reapply);
            return this;
        }

        public IXLFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping, bool reapply)
        {
            SwitchFilter(XLFilterType.Regular);
            AddFilter(XLFilter.CreateDateGroupFilter(date, dateTimeGrouping), reapply);
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
            return AddCustomFilter(value.ToString(), true, reapply);
        }

        public IXLFilterConnector NotEqualTo(XLCellValue value, Boolean reapply)
        {
            return AddCustomFilter(value.ToString(), false, reapply);
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

        /// <summary>
        /// A filter value used by top/bottom filter to compare with cell value.
        /// </summary>
        internal double TopBottomFilterValue { get; private set; } = double.NaN;

        public IEnumerator<XLFilter> GetEnumerator() => _filters.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private void SetTopBottom(Int32 percentOrItemCount, XLTopBottomType type, Boolean takeTop, Boolean reapply)
        {
            if (percentOrItemCount is < 1 or > 500)
                throw new ArgumentOutOfRangeException(nameof(percentOrItemCount), "Value must be between 1 and 500.");

            ResetFilter(XLFilterType.TopBottom);
            TopBottomValue = percentOrItemCount;
            TopBottomType = type;
            TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

            AddFilter(XLFilter.CreateTopBottom(takeTop, percentOrItemCount), reapply);
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

                    // Ceiling, so there is always at least one item.
                    var itemCountByPercents = (int)Math.Ceiling(materializedNumbers.Length * (double)percent / 100);
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

            // `Average` is recalculated during reapply, so no need to calculate it twice.
            DynamicValue = reapply ? double.NaN : GetAverageFilterValue();
            AddFilter(XLFilter.CreateAverage(DynamicValue, aboveAverage), reapply);
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
            AddFilter(XLFilter.CreateCustomFilter(value, op, XLConnector.Or), reapply);
            return new XLFilterConnector(this);
        }

        private IXLFilterConnector AddCustomFilter(string pattern, bool match, bool reapply)
        {
            ResetFilter(XLFilterType.Custom);
            AddFilter(XLFilter.CreateCustomPatternFilter(pattern, match, XLConnector.Or), reapply);
            return new XLFilterConnector(this);
        }

        private void ResetFilter(XLFilterType type)
        {
            Clear(false);
            _autoFilter.IsEnabled = true;
            FilterType = type;
        }

        private void SwitchFilter(XLFilterType type)
        {
            _autoFilter.IsEnabled = true;
            if (FilterType == type)
                return;

            Clear(false);
            FilterType = type;
        }

        internal void AddFilter(XLFilter filter, bool reapply = false)
        {
            var maxFilters = FilterType switch
            {
                XLFilterType.None => 0,
                XLFilterType.Regular => int.MaxValue,
                XLFilterType.Custom => 2,
                XLFilterType.TopBottom => 1,
                XLFilterType.Dynamic => 1,
                _ => throw new NotSupportedException()
            };
            if (_filters.Count >= maxFilters)
                throw new InvalidOperationException($"{FilterType} filter can have max {maxFilters} conditions.");

            _filters.Add(filter);
            if (reapply)
                _autoFilter.Reapply();
        }

        internal void Refresh()
        {
            if (FilterType == XLFilterType.Dynamic)
            {
                // Update average value of a filter, so it is saved correctly and filter uses
                // correct value, even is cell values changed and avg was stale.
                DynamicValue = GetAverageFilterValue();
                _filters[0].Value = DynamicValue;
            }

            if (FilterType == XLFilterType.TopBottom)
            {
                var takeTop = TopBottomPart == XLTopBottomPart.Top;
                TopBottomFilterValue = GetTopBottomFilterValue(TopBottomType, TopBottomValue, takeTop);
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
            var connector = _filters[1].Connector;
            return connector switch
            {
                XLConnector.And => _filters.All(filter => filter.Condition(cell, this)),
                XLConnector.Or => _filters.Any(filter => filter.Condition(cell, this)),
                _ => throw new NotSupportedException(),
            };
        }
    }
}
