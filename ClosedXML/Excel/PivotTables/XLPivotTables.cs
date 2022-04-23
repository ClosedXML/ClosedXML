using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotTables : IXLPivotTables
    {
        private readonly Dictionary<string, XLPivotTable> _pivotTables = new Dictionary<string, XLPivotTable>(StringComparer.OrdinalIgnoreCase);

        public XLPivotTables(IXLWorksheet worksheet)
        {
            this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        internal void Add(string name, IXLPivotTable pivotTable)
        {
            _pivotTables.Add(name, (XLPivotTable)pivotTable);
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLRange range)
        {
            var pivotTable = new XLPivotTable(this.Worksheet)
            {
                Name = name,
                TargetCell = targetCell,
                SourceRange = range
            };
            _pivotTables.Add(name, pivotTable);
            return pivotTable;
        }

        public IXLPivotTable Add(string name, IXLCell targetCell, IXLTable table)
        {
            var pivotTable = new XLPivotTable(this.Worksheet)
            {
                Name = name,
                TargetCell = targetCell,
                SourceTable = table
            };
            _pivotTables.Add(name, pivotTable);
            return pivotTable;
        }

        public IXLPivotTable AddNew(string name, IXLCell targetCell, IXLRange range)
        {
            return Add(name, targetCell, range);
        }

        public IXLPivotTable AddNew(string name, IXLCell targetCell, IXLTable table)
        {
            return Add(name, targetCell, table);
        }

        public bool Contains(string name)
        {
            return _pivotTables.ContainsKey(name);
        }

        public void Delete(string name)
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

        public XLPivotTable PivotTable(string name)
        {
            return _pivotTables[name];
        }

        IXLPivotTable IXLPivotTables.PivotTable(string name)
        {
            return PivotTable(name);
        }

        public IXLWorksheet Worksheet { get; private set; }
    }
}
