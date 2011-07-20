using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLPivotFields: IXLPivotFields
    {
        private readonly Dictionary<String, IXLPivotField> _pivotFields = new Dictionary<string, IXLPivotField>();
        public IEnumerator<IXLPivotField> GetEnumerator()
        {
            return _pivotFields.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLPivotField Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }
        public IXLPivotField Add(String sourceName, String customName)
        {
            var pivotField = new XLPivotField(sourceName) {CustomName = customName};
            _pivotFields.Add(sourceName, pivotField);
            return pivotField;
        }

        public void Clear()
        {
            _pivotFields.Clear();
        }
        public void Remove(String sourceName)
        {
            _pivotFields.Remove(sourceName);
        }
    }
}
