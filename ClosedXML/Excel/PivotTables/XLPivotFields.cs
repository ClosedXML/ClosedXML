using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotFields : IXLPivotFields
    {

        private readonly Dictionary<String, IXLPivotField> _pivotFields = new Dictionary<string, IXLPivotField>();
        private readonly IXLPivotTable _pivotTable;

        internal XLPivotFields(IXLPivotTable pivotTable)
        {
            this._pivotTable = pivotTable;
        }

        public IXLPivotField Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotField Add(String sourceName, String customName)
        {
            if (sourceName != XLConstants.PivotTableValuesSentinalLabel && !this._pivotTable.SourceRangeFieldsAvailable.Contains(sourceName, StringComparer.OrdinalIgnoreCase))
                throw new ArgumentOutOfRangeException(nameof(sourceName), String.Format("The column '{0}' does not appear in the source range.", sourceName));

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
                throw new ArgumentNullException(nameof(pf), "Invalid field name.");

            return selectedItem.Position;
        }

        public void Remove(String sourceName)
        {
            _pivotFields.Remove(sourceName);
        }
    }
}
