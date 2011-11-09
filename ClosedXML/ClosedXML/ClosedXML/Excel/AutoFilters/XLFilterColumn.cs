using System;
using System.Linq;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLFilterColumn: IXLFilterColumn
    {
        private XLAutoFilter _autoFilter;
        private Int32 _column;
        public XLFilterColumn(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        public IXLFilterColumn Sort(XLSortOrder order = XLSortOrder.Ascending) { throw new NotImplementedException(); }
        public void Clear() 
        {
            if (_autoFilter.Filters.ContainsKey(_column))
                _autoFilter.Filters.Remove(_column);
        }
        public IXLFilteredColumn AddFilter<T>(T value) where T : IComparable<T>
        {
            if (typeof(T) == typeof(String))
                ApplyCustomFilter<T>(value, XLFilterOperator.Equal, v => v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase), XLFilterType.Regular);
            else
                ApplyCustomFilter<T>(value, XLFilterOperator.Equal, v => (v.CastTo<T>() as IComparable).CompareTo(value) == 0, XLFilterType.Regular);
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

        private void SetTopBottom(Int32 value, XLTopBottomType type, Boolean takeTop = true)
        {
            _autoFilter.Enabled = true;
            _autoFilter.Column(_column).FilterType = XLFilterType.TopBottom;
            _autoFilter.Column(_column).TopBottomValue = value;
            _autoFilter.Column(_column).TopBottomType = type;
            var column = _autoFilter.Range.Column(_column);
            IEnumerable<Double> values;
            
            if (takeTop)
            {
                if (type == XLTopBottomType.Items)
                    values = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble()).OrderByDescending(d => d).Take(value).Distinct();
                else
                {
                    var numerics = values = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble());
                    Int32 valsToTake = numerics.Count() * value / 100;
                    values = numerics.OrderByDescending(d => d).Take(valsToTake).Distinct();
                }
            }
            else
            {
                if (type == XLTopBottomType.Items)
                    values = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble()).OrderBy(d => d).Take(value).Distinct();
                else
                {
                    var numerics = values = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble());
                    Int32 valsToTake = numerics.Count() * value / 100;
                    values = numerics.OrderBy(d => d).Take(valsToTake).Distinct();
                }
            }

            Clear();
            _autoFilter.Filters.Add(_column, new List<XLFilter>());

            Boolean addToList = true;
            foreach (var row in _autoFilter.Range.Rows().Where(r => r.RowNumber() > 1))
            {
                Boolean foundOne = false;
                foreach (var val in values)
                {
                    Func<Object, Boolean> condition = v => (v as IComparable).CompareTo(val) == 0;
                    if (addToList)
                        _autoFilter.Filters[_column].Add(new XLFilter { Value = val, Operator = XLFilterOperator.Equal, Connector = XLConnector.Or, Condition = condition });

                    var cell = row.Cell(_column);
                    if (cell.DataType == XLCellValues.Number && condition(cell.GetDouble()))
                    {
                        row.WorksheetRow().Unhide();
                        foundOne = true;
                    }
                }
                if (!foundOne)
                    row.WorksheetRow().Hide();

                addToList = false;
                //ApplyCustomFilter<String>(val.ToString(), XLFilterOperator.Equal, s => ((string)s).Equals(val.ToString(), StringComparison.InvariantCultureIgnoreCase));
            }
        }

        public void AboveAverage()
        {
            ShowAverage(true);
        }
        public void BelowAverage()
        {
            ShowAverage(false);
        }

        private void ShowAverage(Boolean aboveAverage)
        {
            _autoFilter.Enabled = true;
            _autoFilter.Column(_column).FilterType = XLFilterType.Dynamic;
            _autoFilter.Column(_column).DynamicType = aboveAverage ? XLFilterDynamicType.AboveAverage : XLFilterDynamicType.BelowAverage;
            var column = _autoFilter.Range.Column(_column);
            Double average = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble()).Average();
            IEnumerable<Double> values;

            if (aboveAverage)
                values = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble()).Where(c => c > average).Distinct();
            else
                values = column.Column(2, column.CellCount()).CellsUsed().Where(c => c.DataType == XLCellValues.Number).Select(c => c.GetDouble()).Where(c => c < average).Distinct();
            

            Clear();
            _autoFilter.Filters.Add(_column, new List<XLFilter>());

            Boolean addToList = true;
            foreach (var row in _autoFilter.Range.Rows().Where(r => r.RowNumber() > 1))
            {
                Boolean foundOne = false;
                foreach (var val in values)
                {
                    Func<Object, Boolean> condition = v => (v as IComparable).CompareTo(val) == 0;
                    if (addToList)
                        _autoFilter.Filters[_column].Add(new XLFilter { Value = val, Operator = XLFilterOperator.Equal, Connector = XLConnector.Or, Condition = condition });

                    var cell = row.Cell(_column);
                    if (cell.DataType == XLCellValues.Number && condition(cell.GetDouble()))
                    {
                        row.WorksheetRow().Unhide();
                        foundOne = true;
                    }
                }
                if (!foundOne)
                    row.WorksheetRow().Hide();

                addToList = false;
                //ApplyCustomFilter<String>(val.ToString(), XLFilterOperator.Equal, s => ((string)s).Equals(val.ToString(), StringComparison.InvariantCultureIgnoreCase));
            }
        }

        public IXLFilterConnector EqualTo<T>(T value) where T : IComparable<T>
        {
            if (typeof(T) == typeof(String))
                return ApplyCustomFilter<T>(value, XLFilterOperator.Equal, v => v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
            else
                return ApplyCustomFilter<T>(value, XLFilterOperator.Equal, v => (v.CastTo<T>() as IComparable).CompareTo(value) == 0);
        }

        public IXLFilterConnector NotEqualTo<T>(T value) where T : IComparable<T>
        {
            //return ApplyCustomFilter<String>(value.ToString(), XLFilterOperator.NotEqual, s => !((string)s).Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));

            if (typeof(T) == typeof(String))
                return ApplyCustomFilter<T>(value, XLFilterOperator.NotEqual, v => !v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
            else
                return ApplyCustomFilter<T>(value, XLFilterOperator.NotEqual, v => (v.CastTo<T>() as IComparable).CompareTo(value) != 0);
        }

        public IXLFilterConnector GreaterThan<T>(T value) where T : IComparable<T>
        {
            return ApplyCustomFilter<T>(value, XLFilterOperator.GreaterThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) > 0);
        }

        public IXLFilterConnector LessThan<T>(T value) where T : IComparable<T>
        {
            return ApplyCustomFilter<T>(value, XLFilterOperator.LessThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) < 0);
        }

        public IXLFilterConnector EqualOrGreaterThan<T>(T value) where T : IComparable<T>
        {
            return ApplyCustomFilter<T>(value, XLFilterOperator.EqualOrGreaterThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) >= 0);
        }

        public IXLFilterConnector EqualOrLessThan<T>(T value) where T : IComparable<T>
        {
            return ApplyCustomFilter<T>(value, XLFilterOperator.EqualOrLessThan, v => (v.CastTo<T>() as IComparable).CompareTo(value) <= 0);
        }

        public void Between<T>(T minValue, T maxValue) where T : IComparable<T>
        {
            EqualOrGreaterThan(minValue).And.EqualOrLessThan(maxValue);
        }
        public void NotBetween<T>(T minValue, T maxValue) where T : IComparable<T>
        {
            LessThan(minValue).Or.GreaterThan(maxValue);
        }

        public IXLFilterConnector BeginsWith(String value) 
        {
            return ApplyCustomFilter<String>(value + "*", XLFilterOperator.Equal, s => ((string)s).StartsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector NotBeginsWith(String value)
        {
            return ApplyCustomFilter<String>(value + "*", XLFilterOperator.NotEqual, s => !((string)s).StartsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector EndsWith(String value)
        {
            return ApplyCustomFilter<String>("*" + value, XLFilterOperator.Equal, s => ((string)s).EndsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector NotEndsWith(String value)
        {
            return ApplyCustomFilter<String>("*" + value, XLFilterOperator.NotEqual, s => !((string)s).EndsWith(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public IXLFilterConnector Contains(String value)
        {
            return ApplyCustomFilter<String>("*" + value + "*", XLFilterOperator.Equal, s => ((string)s).ToLower().Contains(value.ToLower()));
        }
        public IXLFilterConnector NotContains(String value)
        {
            return ApplyCustomFilter<String>("*" + value + "*", XLFilterOperator.Equal, s => !((string)s).ToLower().Contains(value.ToLower()));
        }

        private IXLFilterConnector ApplyCustomFilter<T>(T value, XLFilterOperator op, Func<Object, Boolean> condition, XLFilterType filterType = XLFilterType.Custom) where T : IComparable<T>
        {
            _autoFilter.Enabled = true;
            if (filterType == XLFilterType.Custom)
            {
                Clear();
                _autoFilter.Filters.Add(_column, new List<XLFilter> { new XLFilter { Value = value, Operator = op, Connector = XLConnector.Or, Condition = condition } });
            }
            else
            {
                List<XLFilter> filterList;
                if (_autoFilter.Filters.TryGetValue(_column, out filterList))
                    filterList.Add(new XLFilter { Value = value, Operator = op, Connector = XLConnector.Or, Condition = condition });
                else
                    _autoFilter.Filters.Add(_column, new List<XLFilter> { new XLFilter { Value = value, Operator = op, Connector = XLConnector.Or, Condition = condition } });
            }
            _autoFilter.Column(_column).FilterType = filterType;
            Boolean isText = typeof(T) == typeof(String);
            foreach (var row in _autoFilter.Range.Rows().Where(r => r.RowNumber() > 1))
            {
                Boolean match = isText ? condition(row.Cell(_column).GetString()) : row.Cell(_column).DataType == XLCellValues.Number && condition(row.Cell(_column).GetDouble());
                if (match)
                {
                    row.WorksheetRow().Unhide();
                }
                else
                    row.WorksheetRow().Hide();
            }

            return new XLFilterConnector(_autoFilter, _column);
        }

        public XLFilterType FilterType { get; set; }

        public Int32 TopBottomValue { get; set; }
        public XLTopBottomType TopBottomType { get; set; }
        public XLTopBottomPart TopBottomPart { get; set; }

        public XLFilterDynamicType DynamicType { get; set; }
        public Double DynamicValue { get; set; }
    }
}