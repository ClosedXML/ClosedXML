using System;
using System.Collections.Generic;

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

        internal XLPivotCacheValues(XLPivotCacheSharedItems sharedItems, List<XLPivotCacheValue> values)
        {
            _sharedItems = sharedItems;
            _values = values;
            _stringStorage = new List<string>();
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
            _values.Add(XLPivotCacheValue.ForText(text, _stringStorage));
        }

        internal void AddDateTime(DateTime dateTime)
        {
            _values.Add(XLPivotCacheValue.ForDateTime(dateTime));
        }

        internal void AddIndex(uint index, XLPivotCacheSharedItems sharedItems)
        {
            if (index >= sharedItems.Count)
                throw new ArgumentException("Index is referencing non-existent shared item.");

            _values.Add(XLPivotCacheValue.ForIndex(index));
        }

        internal IEnumerable<XLCellValue> GetCellValues()
        {
            foreach (var value in _values)
            {
                yield return value.GetCellValue(_stringStorage, _sharedItems);
            }
        }
    }
}
