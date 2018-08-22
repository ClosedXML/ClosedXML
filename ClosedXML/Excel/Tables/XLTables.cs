using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLTables : IXLTables
    {
        private readonly Dictionary<String, IXLTable> _tables;

        public XLTables()
        {
            _tables = new Dictionary<String, IXLTable>(StringComparer.OrdinalIgnoreCase);
            Deleted = new HashSet<String>();
        }

        internal ICollection<String> Deleted { get; private set; }

        #region IXLTables Members

        public void Add(IXLTable table)
        {
            _tables.Add(table.Name, table);
            (table as XLTable)?.OnAddedToTables();
        }

        public IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            _tables.Values.ForEach(t => t.Clear(clearOptions));
            return this;
        }

        public Boolean Contains(String name)
        {
            return _tables.ContainsKey(name);
        }

        public IEnumerator<IXLTable> GetEnumerator()
        {
            return _tables.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(Int32 index)
        {
            this.Remove(_tables.ElementAt(index).Key);
        }

        public void Remove(String name)
        {
            if (!_tables.ContainsKey(name))
                throw new ArgumentOutOfRangeException(nameof(name), $"Unable to delete table because the table name {name} could not be found.");

            var table = _tables[name] as XLTable;
            _tables.Remove(name);

            if (table.RelId != null) Deleted.Add(table.RelId);
        }

        public IXLTable Table(Int32 index)
        {
            return _tables.ElementAt(index).Value;
        }

        public IXLTable Table(String name)
        {
            if (TryGetTable(name, out IXLTable table))
                return table;

            throw new ArgumentOutOfRangeException(nameof(name), $"Table {name} was not found.");
        }

        public bool TryGetTable(string tableName, out IXLTable table)
        {
            if (_tables.ContainsKey(tableName))
            {
                table = _tables[tableName];
                return true;
            }
            table = null;
            return false;
        }

        #endregion IXLTables Members
    }
}
