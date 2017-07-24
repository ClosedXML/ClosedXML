using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLPivotValues: IXLPivotValues
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
            var pivotValue = new XLPivotValue(sourceName) { CustomName = customName };
            _pivotValues.Add(customName, pivotValue);

            if (_pivotValues.Count > 1 && !this._pivotTable.ColumnLabels.Any(cl => cl.SourceName == XLConstants.PivotTableValuesSentinalLabel) && !this._pivotTable.RowLabels.Any(rl => rl.SourceName == XLConstants.PivotTableValuesSentinalLabel))
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
