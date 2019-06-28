using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotTables : IXLPivotTables
    {
        private readonly Dictionary<String, XLPivotTable> _pivotTables = new Dictionary<string, XLPivotTable>(StringComparer.OrdinalIgnoreCase);

        public XLPivotTables(IXLWorksheet worksheet)
        {
            this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        internal void Add(String name, IXLPivotTable pivotTable)
        {
            _pivotTables.Add(name, (XLPivotTable)pivotTable);
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLPivotSource pivotSource)
        {
            if (!pivotSource.CachedFields.Any())
                pivotSource.Refresh();

            var pivotTable = new XLPivotTable(this.Worksheet)
            {
                Name = name,
                TargetCell = targetCell,
                Source = pivotSource
            };
            _pivotTables.Add(name, pivotTable);
            return pivotTable;
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLRange range)
        {
            if (!this.Worksheet.Workbook.PivotSources.TryGet(range, out IXLPivotSource pivotSource))
                pivotSource = this.Worksheet.Workbook.PivotSources.Add(range);

            return Add(name, targetCell, pivotSource);
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLTable table)
        {
            if (!this.Worksheet.Workbook.PivotSources.TryGet(table, out IXLPivotSource pivotSource))
                pivotSource = this.Worksheet.Workbook.PivotSources.Add(table);

            return Add(name, targetCell, pivotSource);
        }

        public IXLPivotTable AddNew(string name, IXLCell targetCell, IXLRange range)
        {
            return Add(name, targetCell, range);
        }

        public IXLPivotTable AddNew(string name, IXLCell targetCell, IXLTable table)
        {
            return Add(name, targetCell, table);
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

        public IEnumerator<IXLPivotTable> GetEnumerator()
        {
            return _pivotTables.Values.Cast<IXLPivotTable>().GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public XLPivotTable PivotTable(String name)
        {
            return _pivotTables[name];
        }

        IXLPivotTable IXLPivotTables.PivotTable(String name)
        {
            return PivotTable(name);
        }

        public IXLWorksheet Worksheet { get; private set; }
    }
}
