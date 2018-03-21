using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLTables : IXLTables
    {
        private readonly Dictionary<String, IXLTable> _tables;
        internal ICollection<String> Deleted { get; private set; }

        public XLTables()
        {
            _tables = new Dictionary<String, IXLTable>();
            Deleted = new HashSet<String>();
        }

        #region IXLTables Members

        public IEnumerator<IXLTable> GetEnumerator()
        {
            return _tables.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IXLTable table)
        {
            _tables.Add(table.Name, table);
        }

        public IXLTable Table(Int32 index)
        {
            return _tables.ElementAt(index).Value;
        }

        public IXLTable Table(String name)
        {
            return _tables[name];
        }

        #endregion IXLTables Members

        public IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            _tables.Values.ForEach(t => t.Clear(clearOptions));
            return this;
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
    }
}
