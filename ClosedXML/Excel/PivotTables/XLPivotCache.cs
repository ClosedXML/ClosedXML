using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotCache : IXLPivotCache
    {
        private readonly Dictionary<String, Int32> _fieldIndexes = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<String> _fieldNames = new();
        private readonly List<List<XLCellValue>> _fieldValues = new();

        public XLPivotCache(IXLRange sourceRange)
            : this(new XLPivotSourceReference(sourceRange))
        {
        }

        public XLPivotCache(IXLTable table)
            : this(new XLPivotSourceReference(table))
        {
        }

        private XLPivotCache(XLPivotSourceReference reference)
        {
            Guid = Guid.NewGuid();
            SetExcelDefaults();
            PivotSourceReference = reference;
        }

        #region IXLPivotCache members

        public IReadOnlyList<String> FieldNames => _fieldNames;

        public XLItemsToRetain ItemsToRetainPerField { get; set; }

        public Boolean RefreshDataOnOpen { get; set; }

        public Boolean SaveSourceData { get; set; }

        public IXLPivotCache Refresh()
        {
            _fieldIndexes.Clear();
            _fieldNames.Clear();
            _fieldValues.Clear();

            foreach (var column in PivotSourceReference.SourceRange.Columns())
            {
                var header = column.FirstCell().GetFormattedString();
                var values = column.Cells().Skip(1).Select(c => c.Value).Distinct().ToList();

                AddField(AdjustedFieldName(header), values);
            }

            return this;
        }

        public IXLPivotCache SetItemsToRetainPerField(XLItemsToRetain value) { ItemsToRetainPerField = value; return this; }

        public IXLPivotCache SetRefreshDataOnOpen() => SetRefreshDataOnOpen(true);

        public IXLPivotCache SetRefreshDataOnOpen(Boolean value) { RefreshDataOnOpen = value; return this; }

        public IXLPivotCache SetSaveSourceData() => SetSaveSourceData(true);

        public IXLPivotCache SetSaveSourceData(Boolean value) { SaveSourceData = value; return this; }

        #endregion

        internal IList<String> SourceRangeFields
        {
            get
            {
                // TODO: Once pivot cache is filled with values, replace with fields of a cache.
                return PivotSourceReference.SourceRange
                  .FirstRow()
                  .Cells()
                  .Select(c => c.GetString())
                  .ToList();
            }
        }

        /// <summary>
        /// Pivot cache definition id from the file.
        /// </summary>
        internal uint? CacheId { get; set; }

        internal Guid Guid { get; }

        internal XLPivotSourceReference PivotSourceReference { get; set; }

        internal String? WorkbookCacheRelId { get; set; }

        internal XLPivotCache AddCachedField(String fieldName, List<XLCellValue> items)
        {
            if (_fieldNames.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"Source already contains field {fieldName}.");
            }

            AddField(fieldName, items);
            return this;
        }

        /// <summary>
        /// Try to get a field index for a field name.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="index">The found index, start at 0.</param>
        /// <returns>True if source contains the field.</returns>
        internal bool TryGetFieldIndex(String fieldName, out int index)
        {
            return _fieldIndexes.TryGetValue(fieldName, out index);
        }

        internal bool ContainsField(String fieldName) => _fieldIndexes.ContainsKey(fieldName);
        
        internal IReadOnlyList<XLCellValue> GetFieldValues(String fieldName)
        {
            var index = _fieldIndexes[fieldName];
            return _fieldValues[index];
        }

        internal IList<XLCellValue> GetFieldValues(int fieldIndex)
        {
            return _fieldValues[fieldIndex];
        }

        private String AdjustedFieldName(String header)
        {
            var modifiedHeader = header;
            var i = 1;
            while (_fieldNames.Contains(modifiedHeader, StringComparer.OrdinalIgnoreCase))
            {
                i++;
                modifiedHeader = header + i.ToInvariantString();
            }

            return modifiedHeader;
        }

        private void AddField(String fieldName, List<XLCellValue> items)
        {
            _fieldIndexes.Add(fieldName, _fieldNames.Count);
            _fieldNames.Add(fieldName);
            _fieldValues.Add(items);
        }

        private void SetExcelDefaults()
        {
            SaveSourceData = true;
        }
    }
}
