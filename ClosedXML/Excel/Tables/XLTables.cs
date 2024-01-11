#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLTables : IXLTables, IEnumerable<XLTable>
    {
        private readonly Dictionary<String, XLTable> _tables;

        public XLTables()
        {
            _tables = new Dictionary<String, XLTable>(StringComparer.OrdinalIgnoreCase);
            Deleted = new HashSet<String>();
        }

        internal ICollection<String> Deleted { get; }

        #region IXLTables Members

        bool IXLTables.TryGetTable(string tableName, out IXLTable table)
        {
            if (TryGetTable(tableName, out var foundTable))
            {
                table = foundTable;
                return true;
            }

            table = default;
            return false;
        }

        public void Add(IXLTable table)
        {
            var xlTable = (XLTable)table;
            _tables.Add(table.Name, xlTable);
            xlTable.OnAddedToTables();
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

        public Dictionary<string, XLTable>.ValueCollection.Enumerator GetEnumerator()
        {
            return _tables.Values.GetEnumerator();
        }

        IEnumerator<XLTable> IEnumerable<XLTable>.GetEnumerator() => GetEnumerator();

        IEnumerator<IXLTable> IEnumerable<IXLTable>.GetEnumerator() => GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public void Remove(Int32 index)
        {
            this.Remove(_tables.ElementAt(index).Key);
        }

        public void Remove(String name)
        {
            if (!_tables.TryGetValue(name, out var table))
                throw new ArgumentOutOfRangeException(nameof(name), $"Unable to delete table because the table name {name} could not be found.");

            _tables.Remove(name);

            var relId = table.RelId;

            if (relId is not null)
                Deleted.Add(relId);
        }

        public IXLTable Table(Int32 index)
        {
            return _tables.ElementAt(index).Value;
        }

        public IXLTable Table(String name)
        {
            if (TryGetTable(name, out XLTable table))
                return table;

            throw new ArgumentOutOfRangeException(nameof(name), $"Table {name} was not found.");
        }

        internal bool TryGetTable(string tableName, out XLTable table)
        {
            return _tables.TryGetValue(tableName, out table);
        }

        #endregion IXLTables Members
    }
}
