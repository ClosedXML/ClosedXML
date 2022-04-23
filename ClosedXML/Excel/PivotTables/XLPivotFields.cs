// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotFields : IXLPivotFields
    {
        private readonly Dictionary<string, IXLPivotField> _pivotFields = new Dictionary<string, IXLPivotField>(StringComparer.OrdinalIgnoreCase);
        private readonly IXLPivotTable _pivotTable;

        internal XLPivotFields(IXLPivotTable pivotTable)
        {
            _pivotTable = pivotTable;
        }

        #region IXLPivotFields members

        public IXLPivotField Add(string sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotField Add(string sourceName, string customName)
        {
            if (sourceName != XLConstants.PivotTable.ValuesSentinalLabel && !_pivotTable.SourceRangeFieldsAvailable.Contains(sourceName))
                throw new ArgumentOutOfRangeException(nameof(sourceName), string.Format("The column '{0}' does not appear in the source range.", sourceName));

            var pivotField = new XLPivotField(_pivotTable, sourceName) { CustomName = customName };
            _pivotFields.Add(sourceName, pivotField);
            return pivotField;
        }

        public void Clear()
        {
            _pivotFields.Clear();
        }

        public bool Contains(string sourceName)
        {
            return _pivotFields.ContainsKey(sourceName);
        }

        public bool Contains(IXLPivotField pivotField)
        {
            return _pivotFields.ContainsKey(pivotField.SourceName);
        }

        public IXLPivotField Get(string sourceName)
        {
            return _pivotFields[sourceName];
        }

        public IXLPivotField Get(int index)
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

        public int IndexOf(string sourceName)
        {
            var selectedItem = _pivotFields.Select((item, index) => new { Item = item, Position = index }).FirstOrDefault(i => i.Item.Key == sourceName);
            if (selectedItem == null)
                throw new ArgumentNullException(nameof(sourceName), "Invalid field name.");

            return selectedItem.Position;
        }

        public int IndexOf(IXLPivotField pf)
        {
            return IndexOf(pf.SourceName);
        }

        public void Remove(string sourceName)
        {
            _pivotFields.Remove(sourceName);
        }

        #endregion IXLPivotFields members
    }
}
