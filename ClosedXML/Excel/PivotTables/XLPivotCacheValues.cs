using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Excel.Cells;

namespace ClosedXML.Excel
{
    /// <summary>
    /// All values of a cache field for a pivot table.
    /// </summary>
    internal class XLPivotCacheValues
    {
        private readonly XLPivotCacheSharedItems _sharedItems;

        private readonly List<XLPivotCacheValue> _values;

        private readonly List<string> _stringStorage;

        private bool _containsBlank;

        private bool _containsNumber;

        private double? _minValue;

        private double? _maxValue;

        /// <inheritdoc cref="XLPivotCacheValuesStats.ContainsInteger"/>
        private bool _containsInteger;

        /// <inheritdoc cref="XLPivotCacheValuesStats.ContainsString"/>
        private bool _containsString;

        /// <inheritdoc cref="XLPivotCacheValuesStats.LongText"/>
        private bool _longText;

        /// <inheritdoc cref="XLPivotCacheValuesStats.ContainsDate"/>
        private bool _containsDate;

        private long? _minDateTicks;

        private long? _maxDateTicks;

        internal XLPivotCacheValues(XLPivotCacheSharedItems sharedItems, List<XLPivotCacheValue> values)
        {
            _sharedItems = sharedItems;
            _values = values;
            _stringStorage = new List<string>();
        }

        internal XLPivotCacheValues(XLPivotCacheSharedItems sharedItems, XLPivotCacheValuesStats stats)
        {
            _sharedItems = sharedItems;
            _values = new List<XLPivotCacheValue>();
            _stringStorage = new List<string>();

            // Have a separate fields instead of one large struct. That way,
            // the flags are more easily set when record values are being added.
            _containsBlank = stats.ContainsBlank;
            _containsNumber = stats.ContainsNumber;
            _containsInteger = stats.ContainsInteger;
            _minValue = stats.MinValue;
            _maxValue = stats.MaxValue;
            _containsString = stats.ContainsString;
            _longText = stats.LongText;
            _containsDate = stats.ContainsDate;
            _minDateTicks = stats.MinDate?.Ticks;
            _maxDateTicks = stats.MaxDate?.Ticks;
        }

        internal XLPivotCacheValuesStats Stats
        {
            get
            {
                DateTime? minDate = _containsDate && _minDateTicks is not null ? new DateTime(_minDateTicks.Value) : null;
                DateTime? maxDate = _containsDate && _maxDateTicks is not null ? new DateTime(_maxDateTicks.Value) : null;

                return new XLPivotCacheValuesStats(
                    _containsBlank,
                    _containsNumber,
                    _containsInteger,
                    _minValue,
                    _maxValue,
                    _containsString,
                    _longText,
                    _containsDate,
                    minDate,
                    maxDate);
            }
        }

        internal int Count => _values.Count;

        internal int SharedCount => _sharedItems.Count;

        internal XLPivotCacheSharedItems SharedItems => _sharedItems;

        internal void Add(XLCellValue value)
        {
            switch (value.Type)
            {
                case XLDataType.Blank:
                    AddMissing();
                    break;
                case XLDataType.Boolean:
                    AddBoolean(value.GetBoolean());
                    break;
                case XLDataType.Number:
                    AddNumber(value.GetNumber());
                    break;
                case XLDataType.Text:
                    AddString(value.GetText());
                    break;
                case XLDataType.Error:
                    AddError(value.GetError());
                    break;
                case XLDataType.DateTime:
                    AddDateTime(value.GetDateTime());
                    break;
                case XLDataType.TimeSpan:
                    // TimeSpan is represented as datetime in pivot cache, e.g. 14:30 into 1899-12-30T14:30:00
                    var adjustedTimeSpan = DateTime.FromOADate(0).Add(value.GetTimeSpan());
                    AddDateTime(adjustedTimeSpan);
                    break;
                default:
                    throw new NotSupportedException();
            }

            if (_sharedItems.IndexOf(value) == -1)
                _sharedItems.Add(value);
        }

        internal void AddMissing()
        {
            _values.Add(XLPivotCacheValue.ForMissing());
            _containsBlank = true;
        }

        internal void AddNumber(double number)
        {
            _values.Add(XLPivotCacheValue.ForNumber(number));
            AdjustStats(number);
        }

        internal void AddBoolean(bool boolean)
        {
            _values.Add(XLPivotCacheValue.ForBoolean(boolean));

            // [MS-OI29500]: In Office, boolean and error are considered strings in the context of the containsString attribute.
            _containsString = true;
        }

        internal void AddError(XLError error)
        {
            _values.Add(XLPivotCacheValue.ForError(error));

            // [MS-OI29500]: In Office, boolean and error are considered strings in the context of the containsString attribute.
            _containsString = true;
        }

        internal void AddString(string text)
        {
            _values.Add(XLPivotCacheValue.ForText(text, _stringStorage));
            AdjustStats(text);
        }

        internal void AddDateTime(DateTime dateTime)
        {
            _values.Add(XLPivotCacheValue.ForDateTime(dateTime));
            AdjustStats(dateTime);
        }

        internal void AddIndex(uint index)
        {
            if (index >= _sharedItems.Count)
                throw new ArgumentException("Index is referencing non-existent shared item.");

            _values.Add(XLPivotCacheValue.ForIndex(index));

            // Get value referenced by added index value, so stats can be updated.
            var cacheValue = _sharedItems.GetValue(index);
            switch (cacheValue.Type)
            {
                case XLPivotCacheValueType.Missing:
                    _containsBlank = true;
                    break;
                case XLPivotCacheValueType.Number:
                    AdjustStats(cacheValue.GetNumber());
                    break;
                case XLPivotCacheValueType.Boolean:
                    _containsString = true;
                    break;
                case XLPivotCacheValueType.Error:
                    _containsString = true;
                    break;
                case XLPivotCacheValueType.String:
                    AdjustStats(_sharedItems.GetStringValue(index));
                    break;
                case XLPivotCacheValueType.DateTime:
                    AdjustStats(cacheValue.GetDateTime());
                    break;
                default:
                    throw new NotSupportedException();
            }
        }

        internal XLPivotCacheValue GetValue(int recordIdx)
        {
            return _values[recordIdx];
        }

        internal string GetText(XLPivotCacheValue value)
        {
            Debug.Assert(value.Type == XLPivotCacheValueType.String);
            return value.GetText(_stringStorage);
        }

        internal void AllocateCapacity(int recordCount)
        {
            _values.Capacity = recordCount;
        }

        internal IEnumerable<XLCellValue> GetCellValues()
        {
            foreach (var value in _values)
            {
                yield return value.GetCellValue(_stringStorage, _sharedItems);
            }
        }

        /// <summary>
        /// Get or add a value to the shared items. Throw, if value is not in items.
        /// </summary>
        /// <returns>Index in shared items.</returns>
        internal int GetOrAddSharedItem(XLCellValue value)
        {
            var sharedItemsIndex = _sharedItems.IndexOf(value);
            if (sharedItemsIndex >= 0)
                return sharedItemsIndex;

            // Not in shared items, make sure it actually is present.
            if (!ContainsRecord(value))
                throw new ArgumentException($"Value '{value}' not among cache field values.");

            _sharedItems.Add(value);

            return _sharedItems.Count - 1;
        }

        /// <summary>
        /// Is among the <c>value</c> among values of the record.
        /// </summary>
        private bool ContainsRecord(XLCellValue value)
        {
            for (var index = 0; index < _values.Count; ++index)
            {
                var recordValue = GetValue(index);
                var cacheValue = recordValue.GetCellValue(_stringStorage, _sharedItems);
                if (XLCellValueComparer.OrdinalIgnoreCase.Equals(cacheValue, value))
                    return true;
            }

            return false;
        }

        [SuppressMessage("ReSharper", "CompareOfFloatsByEqualityOperator", Justification = "double.IsInteger() in NET7 uses same method.")]
        private void AdjustStats(double number)
        {
            // containsInt is true only if all numbers are integers.
            _containsInteger =
                // First ever number is an integer.
                (!_containsNumber && number == Math.Truncate(number))
                ||
                // Subsequent number is an integer.
                (_containsInteger && number == Math.Truncate(number));
            _containsNumber = true;
            _minValue = _minValue is null ? number : Math.Min(_minValue.Value, number);
            _maxValue = _maxValue is null ? number : Math.Max(_maxValue.Value, number);
        }

        private void AdjustStats(string text)
        {
            _containsString = true;
            _longText = _longText || text.Length > 255;
        }

        private void AdjustStats(DateTime dateTime)
        {
            _containsDate = true;
            var dateTicks = dateTime.Ticks;
            _minDateTicks = _minDateTicks is null ? dateTicks : Math.Min(_minDateTicks.Value, dateTicks);
            _maxDateTicks = _maxDateTicks is null ? dateTicks : Math.Max(_maxDateTicks.Value, dateTicks);
        }
    }
}
