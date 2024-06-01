using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel.Exceptions;

namespace ClosedXML.Excel
{
    internal class XLPivotCache : IXLPivotCache
    {
        private readonly XLWorkbook _workbook;
        private readonly Dictionary<String, Int32> _fieldIndexes = new(XLHelper.NameComparer);
        private readonly List<String> _fieldNames = new();

        /// <summary>
        /// Length is a number of fields, in same order as <see cref="_fieldNames"/>.
        /// </summary>
        private readonly List<XLPivotCacheValues> _values = new();

        internal XLPivotCache(XLPivotSourceReference reference, XLWorkbook workbook)
        {
            _workbook = workbook;
            Guid = Guid.NewGuid();
            SetExcelDefaults();
            PivotSourceReference = reference;
        }

        #region IXLPivotCache members

        public IReadOnlyList<String> FieldNames => _fieldNames;

        public XLItemsToRetain ItemsToRetainPerField { get; set; }

        public Boolean RefreshDataOnOpen { get; set; }

        public Boolean SaveSourceData { get; set; }

        /// <summary>
        /// Number of fields in the cache.
        /// </summary>
        internal int FieldCount => _fieldNames.Count;

        internal int RecordCount => _fieldNames.Count > 0 ? _values[0].Count : 0;

        public IXLPivotCache Refresh()
        {
            // Refresh can only happen if the reference is valid.
            if (!PivotSourceReference.TryGetSource(_workbook, out var sheet, out var foundArea))
                throw new InvalidReferenceException();

            Debug.Assert(sheet is not null && foundArea is not null);
            var oldFieldNames = _fieldNames.ToList();
            _fieldIndexes.Clear();
            _fieldNames.Clear();
            _values.Clear();

            var valueSlice = sheet.Internals.CellsCollection.ValueSlice;
            var area = foundArea.Value;
            for (var column = area.LeftColumn; column <= area.RightColumn; ++column)
            {
                var header = sheet.Cell(area.TopRow, column).GetFormattedString();
                var sharedItems = new XLPivotCacheSharedItems();

                var fieldRecords = new XLPivotCacheValues(sharedItems, new List<XLPivotCacheValue>());
                for (var row = area.TopRow + 1; row <= area.BottomRow; ++row)
                {
                    var value = valueSlice.GetCellValue(new XLSheetPoint(row, column));
                    switch (value.Type)
                    {
                        case XLDataType.Blank:
                            sharedItems.AddMissing();
                            fieldRecords.AddMissing();
                            break;
                        case XLDataType.Boolean:
                            sharedItems.AddBoolean(value.GetBoolean());
                            fieldRecords.AddBoolean(value.GetBoolean());
                            break;
                        case XLDataType.Number:
                            sharedItems.AddNumber(value.GetNumber());
                            fieldRecords.AddNumber(value.GetNumber());
                            break;
                        case XLDataType.Text:
                            sharedItems.AddString(value.GetText());
                            fieldRecords.AddString(value.GetText());
                            break;
                        case XLDataType.Error:
                            sharedItems.AddError(value.GetError());
                            fieldRecords.AddError(value.GetError());
                            break;
                        case XLDataType.DateTime:
                            sharedItems.AddDateTime(value.GetDateTime());
                            fieldRecords.AddDateTime(value.GetDateTime());
                            break;
                        case XLDataType.TimeSpan:
                            // TimeSpan is represented as datetime in pivot cache, e.g. 14:30 into 1899-12-30T14:30:00
                            var adjustedTimeSpan = DateTime.FromOADate(0).Add(value.GetTimeSpan());
                            fieldRecords.AddDateTime(adjustedTimeSpan);
                            sharedItems.AddDateTime(adjustedTimeSpan);
                            break;
                        default:
                            throw new NotSupportedException();
                    }
                }

                AddField(AdjustedFieldName(header), fieldRecords);
            }

            UpdatePivotTables();
            return this;

            void UpdatePivotTables()
            {
                foreach (var worksheet in _workbook.WorksheetsInternal)
                {
                    foreach (var pivotTable in worksheet.PivotTables)
                    {
                        if (pivotTable.PivotCache == this)
                            pivotTable.UpdateCacheFields(oldFieldNames);
                    }
                }
            }
        }

        public IXLPivotCache SetItemsToRetainPerField(XLItemsToRetain value) { ItemsToRetainPerField = value; return this; }

        public IXLPivotCache SetRefreshDataOnOpen() => SetRefreshDataOnOpen(true);

        public IXLPivotCache SetRefreshDataOnOpen(Boolean value) { RefreshDataOnOpen = value; return this; }

        public IXLPivotCache SetSaveSourceData() => SetSaveSourceData(true);

        public IXLPivotCache SetSaveSourceData(Boolean value) { SaveSourceData = value; return this; }

        #endregion

        /// <summary>
        /// Pivot cache definition id from the file.
        /// </summary>
        internal uint? CacheId { get; set; }

        internal Guid Guid { get; }

        internal XLPivotSourceReference PivotSourceReference { get; set; }

        internal String? WorkbookCacheRelId { get; set; }

        internal XLPivotCache AddCachedField(String fieldName, XLPivotCacheValues fieldValues)
        {
            if (_fieldNames.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"Source already contains field {fieldName}.");
            }

            AddField(fieldName, fieldValues);
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

        internal XLPivotCacheValues GetFieldValues(int fieldIndex)
        {
            return _values[fieldIndex];
        }

        internal XLPivotCacheSharedItems GetFieldSharedItems(int fieldIndex)
        {
            return _values[fieldIndex].SharedItems;
        }

        internal void AllocateRecordCapacity(int recordCount)
        {
            foreach (var fieldValues in _values)
            {
                fieldValues.AllocateCapacity(recordCount);
            }
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

        private void AddField(String fieldName, XLPivotCacheValues fieldValues)
        {
            _fieldIndexes.Add(fieldName, _fieldNames.Count);
            _fieldNames.Add(fieldName);
            _values.Add(fieldValues);
        }

        private void SetExcelDefaults()
        {
            SaveSourceData = true;
        }
    }
}
