// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotFields : IXLPivotFields
    {
        private readonly Dictionary<String, IXLPivotField> _pivotFields = new Dictionary<string, IXLPivotField>(StringComparer.OrdinalIgnoreCase);
        private readonly IXLPivotTable _pivotTable;

        internal XLPivotFields(IXLPivotTable pivotTable)
        {
            this._pivotTable = pivotTable;
        }

        #region IXLPivotFields members

        public IXLPivotField Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotField Add(String sourceName, String customName)
        {
            if (sourceName != XLConstants.PivotTableValuesSentinalLabel && !this._pivotTable.Source.SourceRangeFields.Contains(sourceName))
                throw new ArgumentOutOfRangeException(nameof(sourceName), String.Format("The column '{0}' does not appear in the source range.", sourceName));

            var pivotField = new XLPivotField(_pivotTable, sourceName) { CustomName = customName };
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

        public bool Contains(IXLPivotField pivotField)
        {
            return _pivotFields.ContainsKey(pivotField.SourceName);
        }

        public IXLPivotField Get(String sourceName)
        {
            return _pivotFields[sourceName];
        }

        public IXLPivotField Get(Int32 index)
        {
            return _pivotFields.Values.ElementAt(index);
        }

        public IEnumerator<IXLPivotField> GetEnumerator()
        {
            return _pivotFields.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Int32 IndexOf(String sourceName)
        {
            var selectedItem = _pivotFields.Select((item, index) => new { Item = item, Position = index }).FirstOrDefault(i => i.Item.Key == sourceName);
            if (selectedItem == null)
                throw new ArgumentNullException(nameof(sourceName), "Invalid field name.");

            return selectedItem.Position;
        }

        public Int32 IndexOf(IXLPivotField pf)
        {
            return IndexOf(pf.SourceName);
        }

        public void Remove(String sourceName)
        {
            _pivotFields.Remove(sourceName);
        }

        #endregion IXLPivotFields members
    }
}
