using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotValues : IXLPivotValues
    {
        private readonly Dictionary<String, IXLPivotValue> _pivotValues = new Dictionary<string, IXLPivotValue>();

        private readonly IXLPivotTable _pivotTable;

        internal XLPivotValues(IXLPivotTable pivotTable)
        {
            this._pivotTable = pivotTable;
        }

        public IEnumerator<IXLPivotValue> GetEnumerator()
        {
            return _pivotValues.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLPivotValue Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotValue Add(String sourceName, String customName)
        {
            if (sourceName != XLConstants.PivotTableValuesSentinalLabel && !this._pivotTable.SourceRangeFieldsAvailable.Contains(sourceName, StringComparer.OrdinalIgnoreCase))
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

        public void Remove(String sourceName)
        {
            _pivotValues.Remove(sourceName);
        }
    }
}
