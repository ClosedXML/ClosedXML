// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotTableCalculatedFields : IXLPivotTableCalculatedFields
    {
        private readonly Dictionary<String, IXLPivotTableCalculatedField> _calculatedFields = new Dictionary<String, IXLPivotTableCalculatedField>(StringComparer.OrdinalIgnoreCase);
        private readonly IXLPivotTable _pivotTable;

        internal XLPivotTableCalculatedFields(IXLPivotTable pivotTable)
        {
            this._pivotTable = pivotTable;
        }

        public IXLPivotTableCalculatedField Add(String name, String formula)
        {
            if (_calculatedFields.Keys.Contains(name) || this._pivotTable.SourceRangeFieldsAvailable.Contains(name, StringComparer.OrdinalIgnoreCase))
                throw new ArgumentException(nameof(name), String.Format("The name '{0}' is already in use by another pivot field.", name));

            var calculatedField = new XLPivotTableCalculatedField(name, formula);
            _calculatedFields.Add(name, calculatedField);

            return calculatedField;
        }

        public void Clear()
        {
            foreach (var cf in _calculatedFields.Values)
            {
                var valueFields = _pivotTable.Values
                    .Where(f => f.SourceName.Equals(cf.Name, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var v in valueFields)
                    _pivotTable.Values.Remove(v.CustomName);

                if (_pivotTable.Fields.Contains(cf.Name))
                    _pivotTable.Fields.Remove(cf.Name);
            }

            _calculatedFields.Clear();
        }

        public Boolean Contains(String name)
        {
            return _calculatedFields.ContainsKey(name);
        }

        public IXLPivotTableCalculatedField Get(String name)
        {
            return _calculatedFields[name];
        }

        public IEnumerator<IXLPivotTableCalculatedField> GetEnumerator()
        {
            return _calculatedFields.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(String name)
        {
            _calculatedFields.Remove(name);

            var valueFields = _pivotTable.Values
                .Where(f => f.SourceName.Equals(name, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var v in valueFields)
                _pivotTable.Values.Remove(v.CustomName);

            if (_pivotTable.Fields.Contains(name))
                _pivotTable.Fields.Remove(name);
        }

        public Boolean TryGetCalculatedField(String name, out IXLPivotTableCalculatedField calculatedField)
        {
            return _calculatedFields.TryGetValue(name, out calculatedField);
        }
    }
}
