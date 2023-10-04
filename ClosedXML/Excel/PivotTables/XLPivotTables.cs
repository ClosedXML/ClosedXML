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

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLPivotCache pivotCache)
        {
            if (!pivotCache.FieldNames.Any())
                pivotCache.Refresh();

            var pivotTable = new XLPivotTable(Worksheet)
            {
                Name = name,
                TargetCell = targetCell,
                PivotCache = (XLPivotCache)pivotCache
            };

            _pivotTables.Add(name, pivotTable);
            return pivotTable;
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLRange range)
        {
            var pivotCaches = Worksheet.Workbook.PivotCachesInternal;
            var existingPivotCache = pivotCaches.GetAll(range).FirstOrDefault(s => s.PivotSourceReference.SourceTable is null);
            if (existingPivotCache is null)
            {
                existingPivotCache = pivotCaches.Add(range);
            }

            return Add(name, targetCell, existingPivotCache);
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLTable table)
        {
            var pivotCaches = Worksheet.Workbook.PivotCachesInternal;
            var existingPivotCache = pivotCaches.GetAll(table).FirstOrDefault(s => s.PivotSourceReference.SourceTable is not null);
            if (existingPivotCache is null)
            {
                existingPivotCache = pivotCaches.Add(table);
            }

            return Add(name, targetCell, existingPivotCache);
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
