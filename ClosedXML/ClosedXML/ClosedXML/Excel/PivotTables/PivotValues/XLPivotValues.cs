using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLPivotValues: IXLPivotValues
    {
        private readonly Dictionary<String, IXLPivotValue> _pivotValues = new Dictionary<string, IXLPivotValue>();
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
            _pivotValues.Add(sourceName, pivotValue);
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
