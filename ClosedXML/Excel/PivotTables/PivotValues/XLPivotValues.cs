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
        private readonly Dictionary<String, IXLPivotValue> _pivotValues = new Dictionary<string, IXLPivotValue>(StringComparer.OrdinalIgnoreCase);

        internal XLPivotValues(IXLPivotTable pivotTable)
        {
            this._pivotTable = pivotTable;
        }

        public IXLPivotValue Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotValue Add(String sourceName, String customName)
        {
            if (sourceName != XLConstants.PivotTableValuesSentinalLabel && !this._pivotTable.SourceRangeFieldsAvailable.Contains(sourceName))
                throw new ArgumentOutOfRangeException(nameof(sourceName), String.Format("The column '{0}' does not appear in the source range.", sourceName));

            var pivotValue = new XLPivotValue(sourceName) { CustomName = customName };
            _pivotValues.Add(customName, pivotValue);

            if (_pivotValues.Count > 1 && this._pivotTable.ColumnLabels.All(cl => cl.SourceName != XLConstants.PivotTableValuesSentinalLabel) && this._pivotTable.RowLabels.All(rl => rl.SourceName != XLConstants.PivotTableValuesSentinalLabel))
                _pivotTable.ColumnLabels.Add(XLConstants.PivotTableValuesSentinalLabel);

            return pivotValue;
        }

        public void Clear()
        {
            _pivotValues.Clear();
        }

        public Boolean Contains(String customName)
        {
            return _pivotValues.ContainsKey(customName);
        }

        public Boolean Contains(IXLPivotValue pivotValue)
        {
            return _pivotValues.ContainsKey(pivotValue.SourceName);
        }

        public IXLPivotValue Get(String customName)
        {
            return _pivotValues[customName];
        }

        public IXLPivotValue Get(Int32 index)
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

        public Int32 IndexOf(String customName)
        {
            var selectedItem = _pivotValues.Select((item, index) => new { Item = item, Position = index }).FirstOrDefault(i => i.Item.Key == customName);
            if (selectedItem == null)
                throw new ArgumentNullException(nameof(customName), "Invalid field name.");

            return selectedItem.Position;
        }

        public Int32 IndexOf(IXLPivotValue pivotValue)
        {
            return IndexOf(pivotValue.SourceName);
        }

        public void Remove(String customName)
        {
            _pivotValues.Remove(customName);
        }
    }
}
