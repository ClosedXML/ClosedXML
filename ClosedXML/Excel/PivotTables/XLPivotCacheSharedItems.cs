using System;
using System.Collections.Generic;
using System.Diagnostics;
using ClosedXML.Excel.Cells;

namespace ClosedXML.Excel
{
    /// <summary>
    /// <para>
    /// A list of <see cref="XLPivotCacheValue"/> in the pivot table cache
    /// definition. Generally, it contains all strings of the field records
    /// (record just indexes them through <see cref="XLPivotCacheValueType.Index"/>)
    /// and also values used directly in pivot table (e.g. filter field reference
    /// the table definition, not record).
    /// </para>
    /// <para>
    /// Shared items can't contain <see cref="XLPivotCacheValueType.Index"/>.
    /// </para>
    /// </summary>
    internal class XLPivotCacheSharedItems
    {
        private readonly List<XLPivotCacheValue> _values = new();

        /// <summary>
        /// Storage of strings to save 8 bytes per <c>XLPivotCacheValue</c>
        /// (reference can't be aliased with a number).
        /// </summary>
        private readonly List<string> _stringStorage = new();

        /// <summary>
        /// Strings in a pivot table are case insensitive.
        /// </summary>
        private readonly Dictionary<string, int> _stringMap = new(StringComparer.OrdinalIgnoreCase);

        internal XLCellValue this[uint index] => GetValue(index).GetCellValue(_stringStorage, this);

        internal int Count => _values.Count;

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
                    var timeSpan = value.GetTimeSpan().ToSerialDateTime().ToSerialDateTime();
                    AddDateTime(timeSpan);
                    break;
                default:
                    throw new UnreachableException();
            }
        }

        internal void AddMissing()
        {
            _values.Add(XLPivotCacheValue.ForMissing());
        }

        internal void AddNumber(double number)
        {
            _values.Add(XLPivotCacheValue.ForNumber(number));
        }

        internal void AddBoolean(bool boolean)
        {
            _values.Add(XLPivotCacheValue.ForBoolean(boolean));
        }

        internal void AddError(XLError error)
        {
            _values.Add(XLPivotCacheValue.ForError(error));
        }

        internal void AddString(string text)
        {
            // Shared items doesn't distinguish between two texts that differ only in case.
            if (!_stringMap.ContainsKey(text))
            {
                var index = _stringStorage.Count;
                _values.Add(XLPivotCacheValue.ForText(text, _stringStorage));
                _stringMap.Add(text, index);
            }
        }

        internal void AddDateTime(DateTime dateTime)
        {
            _values.Add(XLPivotCacheValue.ForDateTime(dateTime));
        }

        internal IEnumerable<XLCellValue> GetCellValues()
        {
            foreach (var value in _values)
            {
                yield return value.GetCellValue(_stringStorage, this);
            }
        }

        internal XLPivotCacheValue GetValue(uint index)
        {
            return _values[checked((int)index)];
        }

        internal string GetStringValue(uint index)
        {
            var value = GetValue(index);
            return value.GetText(_stringStorage);
        }

        /// <summary>
        /// Get index of value or -1 if not among shared items.
        /// </summary>
        internal int IndexOf(XLCellValue value)
        {
            for (var index = 0; index < _values.Count; ++index)
            {
                var sharedValue = _values[index];
                var cacheValue = sharedValue.GetCellValue(_stringStorage, this);
                if (XLCellValueComparer.OrdinalIgnoreCase.Equals(cacheValue, value))
                    return index;
            }

            return -1;
        }
    }
}
