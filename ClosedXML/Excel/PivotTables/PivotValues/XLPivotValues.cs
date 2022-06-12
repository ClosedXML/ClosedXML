// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotValues : IXLPivotValues
    {
        private readonly IXLPivotTable _pivotTable;
        private readonly Dictionary<string, IXLPivotValue> _pivotValues = new Dictionary<string, IXLPivotValue>(StringComparer.OrdinalIgnoreCase);

        internal XLPivotValues(IXLPivotTable pivotTable)
        {
            _pivotTable = pivotTable;
        }

        public IXLPivotValue Add(string sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotValue Add(string sourceName, string customName)
        {
            if (sourceName != XLConstants.PivotTable.ValuesSentinalLabel && !_pivotTable.SourceRangeFieldsAvailable.Contains(sourceName))
            {
                throw new ArgumentOutOfRangeException(nameof(sourceName), string.Format("The column '{0}' does not appear in the source range.", sourceName));
            }

            var pivotValue = new XLPivotValue(sourceName) { CustomName = customName };
            _pivotValues.Add(customName, pivotValue);

            if (_pivotValues.Count > 1 && _pivotTable.ColumnLabels.All(cl => cl.SourceName != XLConstants.PivotTable.ValuesSentinalLabel) && _pivotTable.RowLabels.All(rl => rl.SourceName != XLConstants.PivotTable.ValuesSentinalLabel))
            {
                _pivotTable.ColumnLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
            }

            return pivotValue;
        }

        public void Clear()
        {
            _pivotValues.Clear();
        }

        public bool Contains(string sourceName)
        {
            return _pivotValues.ContainsKey(sourceName);
        }

        public bool Contains(IXLPivotValue pivotValue)
        {
            return _pivotValues.ContainsKey(pivotValue.SourceName);
        }

        public IXLPivotValue Get(string sourceName)
        {
            return _pivotValues[sourceName];
        }

        public IXLPivotValue Get(int index)
        {
            return _pivotValues.Values.ElementAt(index);
        }

        public IEnumerator<IXLPivotValue> GetEnumerator()
        {
            return _pivotValues.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int IndexOf(string sourceName)
        {
            var selectedItem = _pivotValues.Select((item, index) => new { Item = item, Position = index }).FirstOrDefault(i => i.Item.Key == sourceName);
            if (selectedItem == null)
            {
                throw new ArgumentNullException(nameof(sourceName), "Invalid field name.");
            }

            return selectedItem.Position;
        }

        public int IndexOf(IXLPivotValue pivotValue)
        {
            return IndexOf(pivotValue.SourceName);
        }

        public void Remove(string sourceName)
        {
            _pivotValues.Remove(sourceName);
        }
    }
}
