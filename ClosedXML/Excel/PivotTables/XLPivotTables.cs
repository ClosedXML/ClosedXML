#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotTables : IXLPivotTables, IEnumerable<XLPivotTable>
    {
        private readonly Dictionary<String, XLPivotTable> _pivotTables = new(StringComparer.OrdinalIgnoreCase);

        public XLPivotTables(XLWorksheet worksheet)
        {
            Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        internal XLWorksheet Worksheet { get; }

        public void Add(XLPivotTable pivotTable)
        {
            var pivotCache = pivotTable.PivotCache;
            if (!pivotCache.FieldNames.Any())
                pivotCache.Refresh();

            _pivotTables.Add(pivotTable.Name, pivotTable);
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLPivotCache pivotCache)
        {
            var pivotTable = new XLPivotTable(Worksheet, (XLPivotCache)pivotCache)
            {
                Name = name,
                TargetCell = targetCell,
                Area = new XLSheetRange(XLSheetPoint.FromAddress(targetCell.Address)),
            };
            Add(pivotTable);
            pivotTable.UpdateCacheFields(Array.Empty<string>());
            return pivotTable;
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLRange range)
        {
            var area = XLBookArea.From(range);
            var pivotCaches = Worksheet.Workbook.PivotCachesInternal;
            var existingPivotCache = pivotCaches.Find(area);
            var pivotCache = existingPivotCache ?? pivotCaches.Add(area);
            return Add(name, targetCell, pivotCache);
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLTable table)
        {
            return Add(name, targetCell, (IXLRange)table);
        }

        public Boolean Contains(String name)
        {
            return _pivotTables.ContainsKey(name);
        }

        public void Delete(String name)
        {
            _pivotTables.Remove(name);
        }

        public void DeleteAll()
        {
            _pivotTables.Clear();
        }

        IXLPivotTable IXLPivotTables.PivotTable(String name)
        {
            return PivotTable(name);
        }

        IEnumerator<IXLPivotTable> IEnumerable<IXLPivotTable>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator<XLPivotTable> IEnumerable<XLPivotTable>.GetEnumerator()
        {
            return GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Dictionary<string, XLPivotTable>.ValueCollection.Enumerator GetEnumerator()
        {
            return _pivotTables.Values.GetEnumerator();
        }

        internal void Add(String name, IXLPivotTable pivotTable)
        {
            _pivotTables.Add(name, (XLPivotTable)pivotTable);
        }

        /// <inheritdoc cref="IXLPivotTables.PivotTable"/>
        internal XLPivotTable PivotTable(String name)
        {
            return _pivotTables[name];
        }
    }
}
