using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    public class XLPivotFields : IXLPivotFields
    {
        private readonly Dictionary<String, IXLPivotField> _pivotFields = new Dictionary<string, IXLPivotField>();

        public IXLPivotField Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotField Add(String sourceName, String customName)
        {
            var pivotField = new XLPivotField(sourceName) { CustomName = customName };
            _pivotFields.Add(sourceName, pivotField);
            return pivotField;
        }

        public void Clear()
        {
            _pivotFields.Clear();
        }

        public Boolean Contains(String sourceName)
        {
            return _pivotFields.ContainsKey(sourceName);
        }

        public IXLPivotField Get(string sourceName)
        {
            return _pivotFields[sourceName];
        }

        public IEnumerator<IXLPivotField> GetEnumerator()
        {
            return _pivotFields.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Int32 IndexOf(IXLPivotField pf)
        {
            var selectedItem = _pivotFields.Select((item, index) => new { Item = item, Position = index }).FirstOrDefault(i => i.Item.Key == pf.SourceName);
            if (selectedItem == null)
                throw new IndexOutOfRangeException("Invalid field name.");
            return selectedItem.Position;
        }

        public void Remove(String sourceName)
        {
            _pivotFields.Remove(sourceName);
        }
    }
}
